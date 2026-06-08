/**
 * Thin Microsoft Graph HTTP client, bound to a single access token.
 *
 * One instance is created per tool invocation (the token is request-scoped), so
 * handlers receive a ready-to-use client via the tool context and never deal
 * with tokens directly. Mirrors the `graphRequest` / `graphUpload` helpers that
 * previously lived in each server, centralized and shared.
 */

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Error thrown when a Graph request fails. `status` is the HTTP status code. */
export class GraphError extends Error {
  readonly status: number;
  /** The machine-readable `error.code` from the Graph error body, if present. */
  readonly graphCode?: string;

  constructor(message: string, status: number, graphCode?: string) {
    super(message);
    this.name = "GraphError";
    this.status = status;
    this.graphCode = graphCode;
  }
}

export type QueryParams = Record<string, string | number | boolean | undefined>;

/** Acceptable raw request body types (a Node-friendly subset of fetch's BodyInit). */
export type GraphBody = string | Uint8Array | ArrayBuffer | Blob;

export interface RawRequestOptions {
  /** JSON body — serialized with JSON.stringify and sent as application/json. */
  body?: unknown;
  /** Raw body (Uint8Array / Buffer / string) sent verbatim; takes precedence over `body`. */
  rawBody?: GraphBody;
  params?: QueryParams;
  headers?: Record<string, string>;
  /**
   * Whether to attach the bearer Authorization header. Defaults to true for
   * relative Graph paths and false for absolute URLs (e.g. pre-authorized
   * upload-session URLs, where sending Authorization can break the request).
   */
  auth?: boolean;
}

const isAbsolute = (path: string): boolean => /^https?:\/\//i.test(path);

const buildUrl = (path: string, params?: QueryParams): string => {
  let url = isAbsolute(path) ? path : `${GRAPH_BASE}${path}`;
  if (params) {
    const search = new URLSearchParams();
    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined) search.set(key, String(value));
    }
    const qs = search.toString();
    if (qs) url += (url.includes("?") ? "&" : "?") + qs;
  }
  return url;
};

export class GraphClient {
  constructor(private readonly accessToken: string) {}

  /** The bound access token. Rarely needed directly. */
  get token(): string {
    return this.accessToken;
  }

  /**
   * Performs a Graph request and returns the parsed JSON body (or null for
   * 204/202 responses). Throws {@link GraphError} on non-2xx responses.
   * Mirrors the original `graphRequest(method, path, token, body?, params?)`.
   */
  async request<T = any>(
    method: string,
    path: string,
    body?: unknown,
    params?: QueryParams,
  ): Promise<T> {
    const res = await this.requestRaw(method, path, { body, params });
    if (res.status === 204 || res.status === 202) return null as T;
    const text = await res.text();
    return (text ? JSON.parse(text) : null) as T;
  }

  /**
   * Uploads binary content (PUT/POST) with an explicit content type and returns
   * the parsed JSON body. Mirrors the original `graphUpload` helper.
   */
  async upload<T = any>(
    method: string,
    path: string,
    content: GraphBody,
    contentType: string,
    params?: QueryParams,
  ): Promise<T> {
    const res = await this.requestRaw(method, path, {
      rawBody: content,
      params,
      headers: { "Content-Type": contentType },
    });
    if (res.status === 204) return null as T;
    const text = await res.text();
    return (text ? JSON.parse(text) : null) as T;
  }

  /**
   * Performs a Graph request and returns the raw `Response` (redirects are
   * followed). Use for downloads, HTTP Range reads, and chunked upload-session
   * PUTs. Throws {@link GraphError} on non-2xx responses; 2xx statuses such as
   * 200/201/202/206 are returned for the caller to inspect.
   */
  async requestRaw(method: string, path: string, opts: RawRequestOptions = {}): Promise<Response> {
    const useAuth = opts.auth ?? !isAbsolute(path);
    const headers: Record<string, string> = {
      Accept: "application/json",
      ...opts.headers,
    };
    if (useAuth) headers["Authorization"] = `Bearer ${this.accessToken}`;
    if (opts.params?.["$search"]) headers["ConsistencyLevel"] = "eventual";

    let body: GraphBody | undefined;
    if (opts.rawBody !== undefined) {
      body = opts.rawBody;
    } else if (opts.body !== undefined) {
      body = JSON.stringify(opts.body);
      if (!headers["Content-Type"]) headers["Content-Type"] = "application/json";
    }

    const res = await fetch(buildUrl(path, opts.params), { method, headers, body });

    if (!res.ok) {
      const errBody: any = await res
        .json()
        .catch(() => ({ error: { message: res.statusText } }));
      // Graph uses `error.message`; the SharePoint REST API (`_api/...`) uses an
      // `odata.error.message.value` envelope — fall back to it when present.
      let message: string =
        errBody?.error?.message ??
        errBody?.["odata.error"]?.message?.value ??
        res.statusText;
      if (res.status === 429) {
        message = "Rate limit exceeded. Please retry after a short delay.";
      }
      throw new GraphError(message, res.status, errBody?.error?.code);
    }

    return res;
  }
}

export const createGraphClient = (accessToken: string): GraphClient => new GraphClient(accessToken);
