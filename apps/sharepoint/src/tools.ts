import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";
import type { GraphClient } from "@microsoft-mcp/graph";

// Files ≤ 4 MB use a direct PUT; larger files use an upload session.
const LARGE_FILE_THRESHOLD = 4 * 1024 * 1024;
// Chunk size must be a multiple of 320 KiB per Graph API requirement.
const CHUNK_SIZE = 20 * 327680; // 6.25 MB

// ---------------------------------------------------------------------------
// Upload helper
// ---------------------------------------------------------------------------

// Upload content in CHUNK_SIZE slices to a pre-created upload session URL.
// The upload URL already has auth embedded — no Authorization header needed.
const uploadInChunks = async (
  graph: GraphClient,
  uploadUrl: string,
  content: Buffer,
  contentType: string,
): Promise<any> => {
  const total = content.length;
  let offset = 0;
  let result: any = null;

  while (offset < total) {
    const end = Math.min(offset + CHUNK_SIZE - 1, total - 1);
    const chunk = content.subarray(offset, end + 1);

    const res = await graph.requestRaw("PUT", uploadUrl, {
      rawBody: new Uint8Array(chunk),
      headers: {
        "Content-Type": contentType,
        "Content-Range": `bytes ${offset}-${end}/${total}`,
        "Content-Length": String(chunk.length),
      },
      auth: false,
    });

    // 202 = more chunks expected; 200/201 = upload complete
    if (res.status === 200 || res.status === 201) {
      result = await res.json();
    }

    offset = end + 1;
  }

  return result;
};

// ---------------------------------------------------------------------------
// Path / URL helpers
// ---------------------------------------------------------------------------

const normalizePath = (p: string): string => (p.startsWith("/") ? p : `/${p}`);

// Returns the Graph path prefix for drive operations, e.g. /sites/{id}/drive
// or /sites/{id}/drives/{driveId}.
const driveBase = (siteId: string, driveId?: string): string =>
  driveId ? `/sites/${siteId}/drives/${driveId}` : `/sites/${siteId}/drive`;

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------

export const tools = [
  // ---- Sites ---------------------------------------------------------------

  defineTool({
    name: "list_sites",
    description: "Search and list SharePoint sites accessible to the signed-in user.",
    inputSchema: {
      search: z
        .string()
        .optional()
        .describe(
          "Keyword to filter sites by display name. Omit or use '*' to list all accessible sites.",
        ),
      top: z
        .number()
        .int()
        .min(1)
        .max(200)
        .optional()
        .describe("Maximum number of sites to return (default 50)."),
      skip: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Number of results to skip for pagination (default 0)."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const query: Record<string, string> = {
        $top: String(params.top ?? 50),
        search: params.search ?? "*",
      };
      if (params.skip) query.$skip = String(params.skip);

      const data = await graph.request("GET", "/sites", undefined, query);
      return data?.value ?? [];
    },
  }),

  defineTool({
    name: "get_site",
    description:
      "Get metadata for a SharePoint site by its Graph site ID or hostname:path form.",
    inputSchema: {
      site_id: z
        .string()
        .describe(
          "Site identifier — either a Graph GUID (e.g. 'contoso.sharepoint.com,abc123,xyz456') or 'hostname:/path' (e.g. 'contoso.sharepoint.com:/sites/hr').",
        ),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      return graph.request("GET", `/sites/${params.site_id}`);
    },
  }),

  defineTool({
    name: "get_site_id",
    description:
      "Resolve a SharePoint site hostname and server-relative path to its Graph GUID. Use the returned id with other tools.",
    inputSchema: {
      hostname: z
        .string()
        .describe("SharePoint hostname, e.g. 'contoso.sharepoint.com'."),
      path: z
        .string()
        .describe(
          "Server-relative site path, e.g. '/sites/hr'. Use '/' for the root site.",
        ),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const path = normalizePath(params.path);
      const data = await graph.request(
        "GET",
        `/sites/${params.hostname}:${path}`,
        undefined,
        { $select: "id,displayName,webUrl" },
      );
      return { id: data.id, displayName: data.displayName, webUrl: data.webUrl };
    },
  }),

  defineTool({
    name: "get_drive",
    description:
      "Get the default document library (drive) for a SharePoint site, or a specific drive when drive_id is supplied.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      drive_id: z
        .string()
        .optional()
        .describe(
          "Specific drive ID. Omit to get the site's default document library.",
        ),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      const endpoint = params.drive_id
        ? `/sites/${params.site_id}/drives/${params.drive_id}`
        : `/sites/${params.site_id}/drive`;
      return graph.request("GET", endpoint);
    },
  }),

  // ---- Lists ---------------------------------------------------------------

  defineTool({
    name: "list_lists",
    description: "List all lists and document libraries in a SharePoint site.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const data = await graph.request(
        "GET",
        `/sites/${params.site_id}/lists`,
        undefined,
        {
          $select:
            "id,name,displayName,description,list,webUrl,createdDateTime,lastModifiedDateTime",
        },
      );
      return data?.value ?? [];
    },
  }),

  defineTool({
    name: "get_list",
    description:
      "Get metadata and column schema for a specific SharePoint list or library.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z
        .string()
        .describe(
          "List ID (GUID) or list display name (e.g. 'Shared Documents').",
        ),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      return graph.request(
        "GET",
        `/sites/${params.site_id}/lists/${params.list_id}`,
        undefined,
        { $expand: "columns" },
      );
    },
  }),

  defineTool({
    name: "list_columns",
    description: "List all column definitions for a SharePoint list.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const data = await graph.request(
        "GET",
        `/sites/${params.site_id}/lists/${params.list_id}/columns`,
      );
      return data?.value ?? [];
    },
  }),

  // ---- List items ----------------------------------------------------------

  defineTool({
    name: "list_items",
    description:
      "Query items in a SharePoint list with optional OData $filter, $select, and pagination.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
      filter: z
        .string()
        .optional()
        .describe(
          "OData $filter expression applied to item fields, e.g. \"fields/Status eq 'Active'\".",
        ),
      select: z
        .string()
        .optional()
        .describe(
          "Comma-separated field names to include in each item's fields object.",
        ),
      top: z
        .number()
        .int()
        .min(1)
        .max(5000)
        .optional()
        .describe("Maximum items to return (default 100)."),
      skip: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Number of items to skip (default 0)."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const query: Record<string, string> = {
        $expand: "fields",
        $top: String(params.top ?? 100),
      };
      if (params.filter) query.$filter = params.filter;
      if (params.select) query["$expand"] = `fields($select=${params.select})`;
      if (params.skip) query.$skip = String(params.skip);

      const data = await graph.request(
        "GET",
        `/sites/${params.site_id}/lists/${params.list_id}/items`,
        undefined,
        query,
      );
      return data?.value ?? [];
    },
  }),

  defineTool({
    name: "get_item",
    description:
      "Get a single SharePoint list item by its ID, including all field values.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
      item_id: z.string().describe("List item ID (numeric string)."),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      return graph.request(
        "GET",
        `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}`,
        undefined,
        { $expand: "fields" },
      );
    },
  }),

  defineTool({
    name: "create_item",
    description: "Create a new item in a SharePoint list.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
      fields: z
        .record(z.string(), z.unknown())
        .describe(
          'Key-value map of field internal names to values, e.g. { "Title": "New item", "Status": "Active" }.',
        ),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      return graph.request(
        "POST",
        `/sites/${params.site_id}/lists/${params.list_id}/items`,
        { fields: params.fields },
      );
    },
  }),

  defineTool({
    name: "update_item",
    description: "Update one or more fields of an existing SharePoint list item.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
      item_id: z.string().describe("List item ID."),
      fields: z
        .record(z.string(), z.unknown())
        .describe("Fields to update as internal-name to value pairs."),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      return graph.request(
        "PATCH",
        `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}/fields`,
        params.fields,
      );
    },
  }),

  defineTool({
    name: "delete_item",
    description: "Permanently delete a SharePoint list item.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      list_id: z.string().describe("List ID or display name."),
      item_id: z.string().describe("List item ID to delete."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      await graph.request(
        "DELETE",
        `/sites/${params.site_id}/lists/${params.list_id}/items/${params.item_id}`,
      );
      return { success: true };
    },
  }),

  // ---- Files ---------------------------------------------------------------

  defineTool({
    name: "upload_file",
    description:
      "Upload a file to a SharePoint document library. Files over 4 MB are automatically chunked via an upload session.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe(
          "Destination path including filename, e.g. '/Shared Documents/reports/Q1.xlsx'.",
        ),
      content_base64: z.string().describe("Base64-encoded file content."),
      content_type: z
        .string()
        .optional()
        .describe("MIME type of the file (default: application/octet-stream)."),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the site's default document library."),
      conflict_behavior: z
        .enum(["replace", "rename", "fail"])
        .optional()
        .describe("How to handle name conflicts (default: replace)."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      const content = Buffer.from(params.content_base64, "base64");
      const contentType = params.content_type ?? "application/octet-stream";
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);
      const conflictBehavior = params.conflict_behavior ?? "replace";

      if (content.length <= LARGE_FILE_THRESHOLD) {
        return graph.upload(
          "PUT",
          `${base}/root:${filePath}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`,
          new Uint8Array(content),
          contentType,
        );
      }

      // Large file: create an upload session, then chunk
      const session = await graph.request(
        "POST",
        `${base}/root:${filePath}:/createUploadSession`,
        { item: { "@microsoft.graph.conflictBehavior": conflictBehavior } },
      );

      return uploadInChunks(graph, session.uploadUrl, content, contentType);
    },
  }),

  defineTool({
    name: "download_file",
    description:
      "Download a file from a SharePoint document library. Returns content as base64.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe("File path in the drive, e.g. '/Shared Documents/report.pdf'."),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);

      const res = await graph.requestRaw(
        "GET",
        `${base}/root:${filePath}:/content`,
      );

      const arrayBuffer = await res.arrayBuffer();
      return {
        content_base64: Buffer.from(arrayBuffer).toString("base64"),
        content_type:
          res.headers.get("content-type") ?? "application/octet-stream",
        size: arrayBuffer.byteLength,
      };
    },
  }),

  defineTool({
    name: "create_folder",
    description: "Create a folder in a SharePoint document library.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      folder_name: z.string().describe("Name of the new folder."),
      parent_path: z
        .string()
        .optional()
        .describe(
          "Parent folder path, e.g. '/Shared Documents'. Defaults to drive root.",
        ),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      const base = driveBase(params.site_id, params.drive_id);
      const endpoint = params.parent_path
        ? `${base}/root:${normalizePath(params.parent_path)}:/children`
        : `${base}/root/children`;

      return graph.request("POST", endpoint, {
        name: params.folder_name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      });
    },
  }),

  defineTool({
    name: "delete_file",
    description:
      "Permanently delete a file or folder from a SharePoint document library.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe(
          "Path of the file or folder to delete, e.g. '/Shared Documents/old-report.pdf'.",
        ),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);
      await graph.request("DELETE", `${base}/root:${filePath}:`);
      return { success: true };
    },
  }),

  defineTool({
    name: "get_file_content",
    description:
      "Retrieve the raw text content of a file in SharePoint. Best for plain-text, CSV, JSON, or Markdown files.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe("File path in the drive, e.g. '/Shared Documents/data.csv'."),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);

      const res = await graph.requestRaw(
        "GET",
        `${base}/root:${filePath}:/content`,
      );

      const content = await res.text();
      return { content };
    },
  }),

  defineTool({
    name: "get_file_metadata",
    description:
      "Get file properties (name, size, MIME type, created/modified dates, web URL) without downloading content.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe("File path in the drive, e.g. '/Shared Documents/report.pdf'."),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);
      return graph.request("GET", `${base}/root:${filePath}:`, undefined, {
        $select:
          "id,name,size,file,folder,createdDateTime,lastModifiedDateTime,webUrl,createdBy,lastModifiedBy",
      });
    },
  }),

  defineTool({
    name: "get_document",
    description:
      "Retrieve both metadata and text content of a SharePoint file in one call. Optimised for AI consumption of documents.",
    inputSchema: {
      site_id: z.string().describe("SharePoint Graph site ID."),
      file_path: z
        .string()
        .describe("File path in the drive, e.g. '/Shared Documents/report.txt'."),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID. Omit to use the default drive."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const filePath = normalizePath(params.file_path);
      const base = driveBase(params.site_id, params.drive_id);

      const [metadata, contentRes] = await Promise.all([
        graph.request("GET", `${base}/root:${filePath}:`, undefined, {
          $select:
            "id,name,size,file,createdDateTime,lastModifiedDateTime,webUrl,createdBy,lastModifiedBy",
        }),
        graph.requestRaw("GET", `${base}/root:${filePath}:/content`),
      ]);

      const text_content = await contentRes.text();
      return { ...metadata, text_content };
    },
  }),

  // ---- Search --------------------------------------------------------------

  defineTool({
    name: "search",
    description:
      "Run a Microsoft Search query across SharePoint content with full control over entity types and result shape.",
    inputSchema: {
      query_string: z
        .string()
        .describe("Search query string. KQL syntax supported."),
      entity_types: z
        .array(z.enum(["driveItem", "listItem", "list", "site", "drive"]))
        .optional()
        .describe(
          "Entity types to search (default: ['driveItem', 'listItem']).",
        ),
      size: z
        .number()
        .int()
        .min(1)
        .max(500)
        .optional()
        .describe("Number of results (default 25)."),
      from: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Offset for pagination (default 0)."),
      fields: z
        .array(z.string())
        .optional()
        .describe("Specific managed properties to include in results."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const body = {
        requests: [
          {
            entityTypes: params.entity_types ?? ["driveItem", "listItem"],
            query: { queryString: params.query_string },
            size: params.size ?? 25,
            from: params.from ?? 0,
            ...(params.fields ? { fields: params.fields } : {}),
          },
        ],
      };
      const data = await graph.request("POST", "/search/query", body);
      return data?.value?.[0]?.hitsContainers ?? [];
    },
  }),

  defineTool({
    name: "search_sharepoint",
    description:
      'Focused full-text search across SharePoint lists and document libraries with ranked results. Supports KQL syntax (e.g. \'budget filetype:xlsx\', \'author:"John Smith"\').',
    inputSchema: {
      query: z.string().describe("Search query. KQL syntax supported."),
      size: z
        .number()
        .int()
        .min(1)
        .max(500)
        .optional()
        .describe("Number of results (default 25)."),
      from: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Pagination offset (default 0)."),
      site_url: z
        .string()
        .optional()
        .describe(
          "Scope the search to a specific SharePoint site by its URL, e.g. 'https://contoso.sharepoint.com/sites/hr'.",
        ),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      // Scope to a specific site by appending a KQL path filter
      const query = params.site_url
        ? `${params.query} path:"${params.site_url}"`
        : params.query;

      const body = {
        requests: [
          {
            entityTypes: ["listItem", "driveItem"],
            query: { queryString: query },
            size: params.size ?? 25,
            from: params.from ?? 0,
            fields: [
              "title",
              "path",
              "url",
              "lastModifiedDateTime",
              "author",
              "fileType",
              "id",
              "siteId",
            ],
          },
        ],
      };

      const data = await graph.request("POST", "/search/query", body);

      const containers: any[] = data?.value?.[0]?.hitsContainers ?? [];
      return containers.flatMap((c) =>
        (c.hits ?? []).map((h: any) => ({
          rank: h.rank,
          score: h.score,
          resource: h.resource,
          summary: h.summary,
        })),
      );
    },
  }),

  // ---- User ----------------------------------------------------------------

  defineTool({
    name: "get_user_profile",
    description:
      "Get the signed-in user's Microsoft 365 profile (id, displayName, mail, jobTitle, department, etc.).",
    inputSchema: {},
    confirmationPolicy: "never",
    handler: ({ graph }) => {
      return graph.request("GET", "/me", undefined, {
        $select:
          "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation",
      });
    },
  }),

  // ---- Site creation -------------------------------------------------------

  defineTool({
    name: "create_site",
    description:
      "Create a new SharePoint site. Team sites are provisioned via a Microsoft 365 Group (requires Group.ReadWrite.All). Communication sites are created via the SharePoint REST API (requires Sites.FullControl.All).",
    inputSchema: {
      display_name: z
        .string()
        .describe("Display name for the new site, e.g. 'Marketing Hub'."),
      site_type: z
        .enum(["team", "communication"])
        .describe(
          "'team' creates a group-connected Team site. 'communication' creates a standalone Communication site.",
        ),
      mail_nickname: z
        .string()
        .optional()
        .describe(
          "Email alias for the M365 group (team sites only). Must be unique across the tenant and contain no spaces, e.g. 'marketing-hub'. Derived from display_name if omitted.",
        ),
      description: z
        .string()
        .optional()
        .describe("Optional description for the site."),
      visibility: z
        .enum(["Private", "Public"])
        .optional()
        .describe(
          "Team site visibility (default: Private). Ignored for communication sites.",
        ),
      owners: z
        .array(z.string())
        .optional()
        .describe(
          "UPNs or user IDs to set as site owners (team sites only), e.g. ['alice@contoso.com'].",
        ),
      hostname: z
        .string()
        .optional()
        .describe(
          "SharePoint hostname required for communication sites, e.g. 'contoso.sharepoint.com'.",
        ),
      site_url_alias: z
        .string()
        .optional()
        .describe(
          "URL alias for the communication site path, e.g. 'marketing-hub' → /sites/marketing-hub. Derived from display_name if omitted.",
        ),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      if (params.site_type === "team") {
        // Derive mail nickname: lowercase, spaces → hyphens, strip non-alphanumeric except hyphens
        const nickname =
          params.mail_nickname ??
          params.display_name
            .toLowerCase()
            .replace(/\s+/g, "-")
            .replace(/[^a-z0-9-]/g, "");

        const groupBody: Record<string, unknown> = {
          displayName: params.display_name,
          description: params.description ?? "",
          mailNickname: nickname,
          groupTypes: ["Unified"],
          mailEnabled: true,
          securityEnabled: false,
          visibility: params.visibility ?? "Private",
        };

        // Resolve owner IDs to Graph user references if provided
        if (params.owners?.length) {
          const ownerRefs = await Promise.all(
            (params.owners as string[]).map(async (upnOrId: string) => {
              // If it looks like a UPN, resolve to object ID first
              if (upnOrId.includes("@")) {
                const u = await graph.request("GET", `/users/${upnOrId}`, undefined, {
                  $select: "id",
                });
                return `https://graph.microsoft.com/v1.0/users/${u.id}`;
              }
              return `https://graph.microsoft.com/v1.0/users/${upnOrId}`;
            }),
          );
          groupBody["owners@odata.bind"] = ownerRefs;
        }

        const group = await graph.request("POST", "/groups", groupBody);

        // Poll up to ~30 s for the SharePoint site to be provisioned
        let site: any = null;
        for (let i = 0; i < 6; i++) {
          try {
            site = await graph.request("GET", `/groups/${group.id}/sites/root`);
            break;
          } catch {
            await new Promise<void>((r) => setTimeout(r, 5000));
          }
        }

        return {
          group: {
            id: group.id,
            displayName: group.displayName,
            mailNickname: group.mailNickname,
            visibility: group.visibility,
          },
          site: site
            ? { id: site.id, displayName: site.displayName, webUrl: site.webUrl }
            : {
                status: "provisioning",
                message:
                  "Site is still being provisioned — retry get_site in a few seconds.",
              },
        };
      }

      // Communication site via SharePoint REST API
      if (!params.hostname) {
        throw new Error("hostname is required to create a communication site.");
      }

      const urlAlias =
        params.site_url_alias ??
        params.display_name
          .toLowerCase()
          .replace(/\s+/g, "-")
          .replace(/[^a-z0-9-]/g, "");

      const spRestUrl = `https://${params.hostname}/_api/SPSiteManager/create`;
      const payload = {
        request: {
          Title: params.display_name,
          Url: `https://${params.hostname}/sites/${urlAlias}`,
          Description: params.description ?? "",
          WebTemplate: "SITEPAGEPUBLISHING#0",
          SiteDesignId: "00000000-0000-0000-0000-000000000000",
          WebTemplateExtensionId: "00000000-0000-0000-0000-000000000000",
          Classification: "",
        },
      };

      const res = await graph.requestRaw("POST", spRestUrl, {
        body: payload,
        headers: {
          Accept: "application/json;odata.metadata=none",
          "Content-Type": "application/json",
          "OData-Version": "4.0",
        },
        auth: true,
      });

      const data: any = await res.json();
      return {
        siteUrl: data.SiteUrl ?? `https://${params.hostname}/sites/${urlAlias}`,
        siteStatus: data.SiteStatus,
        title: params.display_name,
      };
    },
  }),
];
