import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { GraphClient, GraphError, createGraphClient, GRAPH_BASE } from "./index.js";

const fetchMock = vi.fn();

const TOKEN = "test-token-123";

/** Build a JSON Response. */
const jsonResponse = (obj: unknown, status = 200): Response =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { "content-type": "application/json" },
  });

/** Build an empty-body Response (e.g. 204/202). */
const emptyResponse = (status: number): Response =>
  new Response(null, { status });

/** Convenience accessors for the most recent fetch call. */
const lastCall = () => fetchMock.mock.calls[fetchMock.mock.calls.length - 1];
const lastUrl = (): string => lastCall()[0] as string;
const lastInit = (): RequestInit => lastCall()[1] as RequestInit;
const lastHeaders = (): Record<string, string> =>
  (lastInit().headers ?? {}) as Record<string, string>;

beforeEach(() => {
  vi.stubGlobal("fetch", fetchMock);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

describe("GRAPH_BASE", () => {
  it("is the v1.0 Graph endpoint", () => {
    expect(GRAPH_BASE).toBe("https://graph.microsoft.com/v1.0");
  });
});

describe("createGraphClient", () => {
  it("returns a GraphClient instance", () => {
    const client = createGraphClient(TOKEN);
    expect(client).toBeInstanceOf(GraphClient);
  });

  it("binds the token, exposed via the token getter", () => {
    const client = createGraphClient("abc");
    expect(client.token).toBe("abc");
  });
});

describe("GraphError", () => {
  it("is an instanceof Error with name GraphError", () => {
    const err = new GraphError("boom", 500, "someCode");
    expect(err).toBeInstanceOf(Error);
    expect(err.name).toBe("GraphError");
    expect(err.message).toBe("boom");
    expect(err.status).toBe(500);
    expect(err.graphCode).toBe("someCode");
  });

  it("allows graphCode to be omitted", () => {
    const err = new GraphError("boom", 400);
    expect(err.graphCode).toBeUndefined();
  });
});

describe("URL building", () => {
  beforeEach(() => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
  });

  it("prefixes a relative path with GRAPH_BASE", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events");
    expect(lastUrl()).toBe(`${GRAPH_BASE}/me/events`);
  });

  it("uses an absolute http(s) url as-is", async () => {
    const client = createGraphClient(TOKEN);
    const abs = "https://example.com/upload/session/xyz";
    await client.requestRaw("PUT", abs);
    expect(lastUrl()).toBe(abs);
  });

  it("appends query params with a leading ? when no existing query string", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events", undefined, { top: 5 });
    expect(lastUrl()).toBe(`${GRAPH_BASE}/me/events?top=5`);
  });

  it("appends query params with & when the url already has a query string", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events?foo=bar", undefined, { top: 5 });
    expect(lastUrl()).toBe(`${GRAPH_BASE}/me/events?foo=bar&top=5`);
  });

  it("omits params whose value is undefined", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events", undefined, {
      a: "x",
      b: undefined,
      c: "y",
    });
    const url = lastUrl();
    expect(url).toContain("a=x");
    expect(url).toContain("c=y");
    expect(url).not.toContain("b=");
  });

  it("stringifies number and boolean param values", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events", undefined, {
      n: 42,
      flag: true,
      off: false,
    });
    const url = lastUrl();
    expect(url).toContain("n=42");
    expect(url).toContain("flag=true");
    expect(url).toContain("off=false");
  });

  it("does not append a query string when params object yields no entries", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/events", undefined, { a: undefined });
    expect(lastUrl()).toBe(`${GRAPH_BASE}/me/events`);
  });
});

describe("auth header", () => {
  beforeEach(() => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
  });

  it("sends Authorization: Bearer <token> for a relative path by default", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me");
    expect(lastHeaders()["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });

  it("does NOT send Authorization for an absolute url by default", async () => {
    const client = createGraphClient(TOKEN);
    await client.requestRaw("GET", "https://example.com/thing");
    expect(lastHeaders()["Authorization"]).toBeUndefined();
  });

  it("auth:false on a relative path omits the Authorization header", async () => {
    const client = createGraphClient(TOKEN);
    await client.requestRaw("GET", "/me", { auth: false });
    expect(lastHeaders()["Authorization"]).toBeUndefined();
  });

  it("auth:true on an absolute url adds the Authorization header", async () => {
    const client = createGraphClient(TOKEN);
    await client.requestRaw("GET", "https://example.com/thing", { auth: true });
    expect(lastHeaders()["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });

  it("always includes Accept: application/json", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me");
    expect(lastHeaders()["Accept"]).toBe("application/json");

    await client.requestRaw("GET", "https://example.com/thing");
    expect(lastHeaders()["Accept"]).toBe("application/json");
  });
});

describe("$search consistency header", () => {
  beforeEach(() => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
  });

  it("adds ConsistencyLevel: eventual when a $search param is present", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/messages", undefined, {
      $search: '"hello"',
    });
    expect(lastHeaders()["ConsistencyLevel"]).toBe("eventual");
  });

  it("does not add ConsistencyLevel when no $search param", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me/messages", undefined, { top: 5 });
    expect(lastHeaders()["ConsistencyLevel"]).toBeUndefined();
  });
});

describe("JSON body handling", () => {
  beforeEach(() => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
  });

  it("serializes body with JSON.stringify and sets Content-Type: application/json", async () => {
    const client = createGraphClient(TOKEN);
    const payload = { subject: "Meeting", attendees: [1, 2] };
    await client.request("POST", "/me/events", payload);
    expect(lastInit().body).toBe(JSON.stringify(payload));
    expect(lastHeaders()["Content-Type"]).toBe("application/json");
  });

  it("does not override a caller-supplied Content-Type for a JSON body", async () => {
    const client = createGraphClient(TOKEN);
    await client.requestRaw("POST", "/me/events", {
      body: { a: 1 },
      headers: { "Content-Type": "application/custom" },
    });
    expect(lastInit().body).toBe(JSON.stringify({ a: 1 }));
    expect(lastHeaders()["Content-Type"]).toBe("application/custom");
  });

  it("sends rawBody verbatim and it takes precedence over body", async () => {
    const client = createGraphClient(TOKEN);
    const raw = "verbatim-content";
    await client.requestRaw("PUT", "/me/drive/items/1/content", {
      rawBody: raw,
      body: { ignored: true },
    });
    expect(lastInit().body).toBe(raw);
    // No JSON Content-Type implicitly added for rawBody.
    expect(lastHeaders()["Content-Type"]).toBeUndefined();
  });

  it("sends no body when neither body nor rawBody provided", async () => {
    const client = createGraphClient(TOKEN);
    await client.request("GET", "/me");
    expect(lastInit().body).toBeUndefined();
  });
});

describe("request() return values", () => {
  it("returns parsed JSON on 200", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ id: "1", name: "Alice" }));
    const client = createGraphClient(TOKEN);
    const result = await client.request("GET", "/me");
    expect(result).toEqual({ id: "1", name: "Alice" });
  });

  it("returns null on 204", async () => {
    fetchMock.mockResolvedValue(emptyResponse(204));
    const client = createGraphClient(TOKEN);
    const result = await client.request("DELETE", "/me/events/1");
    expect(result).toBeNull();
  });

  it("returns null on 202", async () => {
    fetchMock.mockResolvedValue(emptyResponse(202));
    const client = createGraphClient(TOKEN);
    const result = await client.request("POST", "/me/sendMail");
    expect(result).toBeNull();
  });

  it("returns null on a 200 with an empty body", async () => {
    fetchMock.mockResolvedValue(new Response("", { status: 200 }));
    const client = createGraphClient(TOKEN);
    const result = await client.request("GET", "/me");
    expect(result).toBeNull();
  });

  it("passes method, path and params through to fetch", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
    const client = createGraphClient(TOKEN);
    await client.request("PATCH", "/me/events/1", { subject: "x" }, { top: 1 });
    expect(lastInit().method).toBe("PATCH");
    expect(lastUrl()).toBe(`${GRAPH_BASE}/me/events/1?top=1`);
  });
});

describe("error handling (requestRaw on non-2xx)", () => {
  it("throws GraphError with status and graphCode from error.code", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse(
        { error: { code: "ErrorItemNotFound", message: "Item not found" } },
        404,
      ),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/me/events/missing")).rejects.toMatchObject({
      name: "GraphError",
      status: 404,
      graphCode: "ErrorItemNotFound",
      message: "Item not found",
    });
  });

  it("throws a real GraphError instance", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({ error: { code: "BadRequest", message: "bad" } }, 400),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toBeInstanceOf(GraphError);
  });

  it("prefers error.message for the message", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({ error: { message: "Primary message" } }, 400),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toThrow("Primary message");
  });

  it("falls back to odata.error.message.value (SharePoint-style envelope)", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse(
        { "odata.error": { message: { value: "SharePoint failure" } } },
        400,
      ),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/_api/web")).rejects.toMatchObject({
      message: "SharePoint failure",
      status: 400,
    });
  });

  it("falls back to statusText when the error body has no recognizable message", async () => {
    fetchMock.mockResolvedValue(
      new Response(JSON.stringify({ something: "else" }), {
        status: 503,
        statusText: "Service Unavailable",
        headers: { "content-type": "application/json" },
      }),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toMatchObject({
      message: "Service Unavailable",
      status: 503,
    });
  });

  it("falls back to statusText when the error body is not valid JSON", async () => {
    fetchMock.mockResolvedValue(
      new Response("not json at all", {
        status: 500,
        statusText: "Internal Server Error",
        headers: { "content-type": "text/plain" },
      }),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toMatchObject({
      message: "Internal Server Error",
      status: 500,
    });
  });

  it("uses the fixed rate-limit message and status 429 on 429", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse(
        { error: { code: "TooManyRequests", message: "throttled" } },
        429,
      ),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toMatchObject({
      message: "Rate limit exceeded. Please retry after a short delay.",
      status: 429,
      graphCode: "TooManyRequests",
    });
  });

  it("error.code is carried even when message comes from a fallback", async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({ error: { code: "SomeCode" } }, 400),
    );
    const client = createGraphClient(TOKEN);
    await expect(client.request("GET", "/x")).rejects.toMatchObject({
      graphCode: "SomeCode",
    });
  });
});

describe("upload()", () => {
  it("sends rawBody with the given Content-Type and returns parsed JSON", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ id: "item-1" }));
    const client = createGraphClient(TOKEN);
    const content = new Uint8Array([1, 2, 3]);
    const result = await client.upload(
      "PUT",
      "/me/drive/root:/file.bin:/content",
      content,
      "application/octet-stream",
    );
    expect(result).toEqual({ id: "item-1" });
    expect(lastInit().method).toBe("PUT");
    expect(lastInit().body).toBe(content);
    expect(lastHeaders()["Content-Type"]).toBe("application/octet-stream");
  });

  it("returns null on 204", async () => {
    fetchMock.mockResolvedValue(emptyResponse(204));
    const client = createGraphClient(TOKEN);
    const result = await client.upload(
      "PUT",
      "/me/drive/root:/file.bin:/content",
      "data",
      "text/plain",
    );
    expect(result).toBeNull();
  });

  it("forwards query params", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
    const client = createGraphClient(TOKEN);
    await client.upload("POST", "/upload", "data", "text/plain", {
      "@microsoft.graph.conflictBehavior": "replace",
    });
    expect(lastUrl()).toContain("conflictBehavior=replace");
  });

  it("attaches the Authorization header for relative upload paths", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ ok: true }));
    const client = createGraphClient(TOKEN);
    await client.upload("PUT", "/me/drive/content", "data", "text/plain");
    expect(lastHeaders()["Authorization"]).toBe(`Bearer ${TOKEN}`);
  });
});

describe("requestRaw() returns the raw Response on 2xx", () => {
  it("returns the Response without consuming it for 200", async () => {
    fetchMock.mockResolvedValue(jsonResponse({ a: 1 }, 200));
    const client = createGraphClient(TOKEN);
    const res = await client.requestRaw("GET", "/me/drive/items/1/content");
    expect(res).toBeInstanceOf(Response);
    expect(res.status).toBe(200);
    await expect(res.json()).resolves.toEqual({ a: 1 });
  });

  it("returns 206 partial-content responses for the caller to inspect", async () => {
    fetchMock.mockResolvedValue(new Response("chunk", { status: 206 }));
    const client = createGraphClient(TOKEN);
    const res = await client.requestRaw("GET", "/me/drive/items/1/content", {
      headers: { Range: "bytes=0-4" },
    });
    expect(res.status).toBe(206);
  });
});
