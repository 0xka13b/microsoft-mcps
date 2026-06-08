import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";
import { escapeODataString, validateDrivePath, validateId } from "@microsoft-mcp/validation";
import { GRAPH_BASE } from "@microsoft-mcp/graph";

export const tools = [
  defineTool({
    name: "me",
    description: "Get the signed-in user's profile (id, displayName, mail).",
    inputSchema: {},
    confirmationPolicy: "never",
    handler: ({ graph }) =>
      graph.request("GET", "/me", undefined, { $select: "id,displayName,mail" }),
  }),

  defineTool({
    name: "list_files",
    description: "List files and folders in OneDrive at a given path.",
    inputSchema: {
      path: z
        .string()
        .optional()
        .describe("Path in OneDrive, e.g. '/' for root or '/Documents'. Defaults to root."),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max items to return. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const path = validateDrivePath(params.path ?? "/", "path");
      const endpoint =
        path === "/"
          ? "/me/drive/root/children"
          : `/me/drive/root:${path}:/children`;

      const data = await graph.request("GET", endpoint, undefined, {
        $top: String(params.limit ?? 50),
        $select:
          "id,name,size,folder,file,createdDateTime,lastModifiedDateTime,webUrl",
      });
      return data.value ?? [];
    },
  }),

  defineTool({
    name: "get_file",
    description:
      "Get metadata for a OneDrive item by ID. Includes a temporary download URL for files.",
    inputSchema: {
      file_id: z.string().describe("Drive item ID"),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      validateId(params.file_id, "file_id");
      const data = await graph.request(
        "GET",
        `/me/drive/items/${params.file_id}`,
        undefined,
        {
          $select:
            "id,name,size,folder,file,createdDateTime,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl",
          $expand: "thumbnails",
        },
      );
      if (data?.["@microsoft.graph.downloadUrl"]) {
        data.download_url = data["@microsoft.graph.downloadUrl"];
        delete data["@microsoft.graph.downloadUrl"];
      } else if (data?.file) {
        const redirectRes = await fetch(
          `${GRAPH_BASE}/me/drive/items/${params.file_id}/content`,
          { headers: { Authorization: `Bearer ${graph.token}` }, redirect: "manual" },
        );
        const cdnUrl = redirectRes.headers.get("location");
        await redirectRes.body?.cancel?.();
        if (cdnUrl) data.download_url = cdnUrl;
      }
      return data;
    },
  }),

  defineTool({
    name: "create_file",
    description: "Upload a new file to OneDrive.",
    inputSchema: {
      onedrive_path: z
        .string()
        .describe("Target path in OneDrive including filename, e.g. '/Documents/report.pdf'"),
      content_base64: z.string().describe("Base64-encoded file content"),
      content_type: z
        .string()
        .optional()
        .describe("MIME type. Defaults to application/octet-stream."),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      const content = Buffer.from(params.content_base64, "base64");
      const raw = params.onedrive_path.startsWith("/")
        ? params.onedrive_path
        : `/${params.onedrive_path}`;
      const path = validateDrivePath(raw, "onedrive_path");

      return graph.upload(
        "PUT",
        `/me/drive/root:${path}:/content`,
        content,
        params.content_type ?? "application/octet-stream",
      );
    },
  }),

  defineTool({
    name: "update_file",
    description: "Replace the content of an existing OneDrive file.",
    inputSchema: {
      file_id: z.string().describe("Drive item ID to update"),
      content_base64: z.string().describe("Base64-encoded new file content"),
      content_type: z
        .string()
        .optional()
        .describe("MIME type. Defaults to application/octet-stream."),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      validateId(params.file_id, "file_id");
      const content = Buffer.from(params.content_base64, "base64");

      return graph.upload(
        "PUT",
        `/me/drive/items/${params.file_id}/content`,
        content,
        params.content_type ?? "application/octet-stream",
      );
    },
  }),

  defineTool({
    name: "delete_file",
    description: "Delete a file or folder from OneDrive.",
    inputSchema: {
      file_id: z.string().describe("Drive item ID to delete"),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      validateId(params.file_id, "file_id");
      await graph.request("DELETE", `/me/drive/items/${params.file_id}`);
      return { success: true };
    },
  }),

  defineTool({
    name: "search_files",
    description: "Search for files in OneDrive by name or content.",
    inputSchema: {
      query: z.string().describe("Search query string"),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max results. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      // Escape single quotes per OData string literal rules before URI encoding
      const q = encodeURIComponent(escapeODataString(params.query));
      const data = await graph.request(
        "GET",
        `/me/drive/root/search(q='${q}')`,
        undefined,
        {
          $top: String(params.limit ?? 50),
          $select:
            "id,name,size,folder,file,createdDateTime,lastModifiedDateTime,webUrl",
        },
      );
      return data.value ?? [];
    },
  }),

  defineTool({
    name: "download_file",
    description:
      "Download a file from OneDrive. Returns the full file content as base64. " +
      "Use ONLY for files under 10 MB — for larger files use download_file_stream.",
    inputSchema: {
      file_id: z.string().describe("Drive item ID to download"),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      validateId(params.file_id, "file_id");

      // Graph returns 302 to a CDN URL; fetch manually to avoid sending auth to CDN
      const redirectRes = await fetch(
        `${GRAPH_BASE}/me/drive/items/${params.file_id}/content`,
        { headers: { Authorization: `Bearer ${graph.token}` }, redirect: "manual" },
      );

      const downloadUrl = redirectRes.headers.get("location");

      if (!downloadUrl && !redirectRes.ok) {
        const err: any = await redirectRes
          .json()
          .catch(() => ({ error: { message: redirectRes.statusText } }));
        throw Object.assign(new Error(err?.error?.message ?? redirectRes.statusText), {
          status: redirectRes.status,
        });
      }

      let arrayBuffer: ArrayBuffer;
      let contentType: string;

      if (downloadUrl) {
        await redirectRes.body?.cancel?.();
        const res = await fetch(downloadUrl);
        if (!res.ok) {
          throw Object.assign(new Error(`CDN download failed: ${res.statusText}`), {
            status: res.status,
          });
        }
        arrayBuffer = await res.arrayBuffer();
        contentType = res.headers.get("content-type") ?? "application/octet-stream";
      } else {
        arrayBuffer = await redirectRes.arrayBuffer();
        contentType = redirectRes.headers.get("content-type") ?? "application/octet-stream";
      }

      return {
        content_base64: Buffer.from(arrayBuffer).toString("base64"),
        content_type: contentType,
        size: arrayBuffer.byteLength,
      };
    },
  }),

  defineTool({
    name: "download_file_stream",
    description:
      "Download a file from OneDrive in chunks using HTTP Range requests. " +
      "Use for files 10 MB or larger. Call repeatedly with increasing offset values until has_more is false. " +
      "Returns content_base64, content_type, offset, chunk_size, total_size, and has_more.",
    inputSchema: {
      file_id: z.string().describe("Drive item ID to download"),
      offset: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Byte offset to start reading from. Defaults to 0."),
      chunk_size: z
        .number()
        .int()
        .min(1)
        .max(10 * 1024 * 1024)
        .optional()
        .describe("Bytes to read per call. Defaults to 5 MB, max 10 MB."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      validateId(params.file_id, "file_id");

      const offset = params.offset ?? 0;
      const chunkSize = params.chunk_size ?? 5 * 1024 * 1024;

      // Resolve CDN URL via redirect
      const redirectRes = await fetch(
        `${GRAPH_BASE}/me/drive/items/${params.file_id}/content`,
        { headers: { Authorization: `Bearer ${graph.token}` }, redirect: "manual" },
      );

      const downloadUrl = redirectRes.headers.get("location");
      if (!downloadUrl && !redirectRes.ok) {
        const err: any = await redirectRes
          .json()
          .catch(() => ({ error: { message: redirectRes.statusText } }));
        throw Object.assign(new Error(err?.error?.message ?? redirectRes.statusText), {
          status: redirectRes.status,
        });
      }
      if (!downloadUrl) {
        await redirectRes.body?.cancel?.();
        throw Object.assign(new Error("No redirect location in 302 response"), {
          status: redirectRes.status,
        });
      }
      await redirectRes.body?.cancel?.();

      const rangeEnd = offset + chunkSize - 1;
      const res = await fetch(downloadUrl, {
        headers: { Range: `bytes=${offset}-${rangeEnd}` },
      });

      if (!res.ok && res.status !== 206) {
        throw Object.assign(new Error(`CDN download failed: ${res.statusText}`), {
          status: res.status,
        });
      }

      const arrayBuffer = await res.arrayBuffer();
      const contentType = res.headers.get("content-type") ?? "application/octet-stream";

      // Parse total size from Content-Range header (bytes 0-4999/1234567)
      const contentRange = res.headers.get("content-range");
      const totalPart = contentRange ? contentRange.split("/")[1] : undefined;
      const parsedTotal = totalPart ? parseInt(totalPart, 10) : NaN;
      const totalSize = Number.isNaN(parsedTotal)
        ? offset + arrayBuffer.byteLength
        : parsedTotal;

      const receivedBytes = arrayBuffer.byteLength;

      return {
        content_base64: Buffer.from(arrayBuffer).toString("base64"),
        content_type: contentType,
        offset,
        chunk_size: receivedBytes,
        total_size: totalSize,
        has_more: offset + receivedBytes < totalSize,
      };
    },
  }),
];
