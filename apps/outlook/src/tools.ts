import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";
import { GraphError } from "@microsoft-mcp/graph";
import { escapeKql, httpError, validateId } from "@microsoft-mcp/validation";

// Replicates the original behavior: a 404 from a message endpoint surfaces a
// friendlier "Email not found" message instead of the raw Graph error.
const emailReq = async <T>(p: Promise<T>): Promise<T> => {
  try {
    return await p;
  } catch (err) {
    if (err instanceof GraphError && err.status === 404) {
      throw new GraphError(
        "Email not found. It may have been deleted or the ID is incorrect.",
        404,
        err.graphCode,
      );
    }
    throw err;
  }
};

const FOLDER_MAP: Record<string, string> = {
  inbox: "inbox",
  sent: "sentitems",
  drafts: "drafts",
  deleted: "deleteditems",
  junk: "junkemail",
  archive: "archive",
};

const resolveFolder = (folder: string): string => {
  const mapped = FOLDER_MAP[folder.toLowerCase()];
  if (mapped) return mapped;
  if (/[/?#]/.test(folder)) throw httpError("Invalid folder identifier", 400);
  return folder;
};

const toRecipients = (emails: string[]) =>
  emails.map((addr) => ({ emailAddress: { address: addr } }));

const toAttachments = (
  attachments?: Array<{
    name: string;
    contentType: string;
    contentBytes: string;
  }>,
) =>
  attachments?.map((a) => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: a.name,
    contentType: a.contentType,
    contentBytes: a.contentBytes,
  }));

const buildEmailBody = (params: {
  to: string[];
  subject: string;
  body: string;
  body_type?: string;
  cc?: string[];
  bcc?: string[];
  attachments?: Array<{ name: string; contentType: string; contentBytes: string }>;
}) => {
  const msg: any = {
    subject: params.subject,
    body: {
      contentType: params.body_type === "html" ? "HTML" : "Text",
      content: params.body,
    },
    toRecipients: toRecipients(params.to),
  };
  if (params.cc?.length) msg.ccRecipients = toRecipients(params.cc);
  if (params.bcc?.length) msg.bccRecipients = toRecipients(params.bcc);
  if (params.attachments?.length) msg.attachments = toAttachments(params.attachments);
  return msg;
};

const recipient = z.string().describe("Email address");

const attachment = z.object({
  name: z.string().describe("Filename"),
  contentType: z.string().describe("MIME type, e.g. 'application/pdf'"),
  contentBytes: z.string().describe("Base64-encoded file content"),
});

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
    name: "list_emails",
    description: "List emails from a mailbox folder ordered by date descending.",
    inputSchema: {
      folder: z
        .string()
        .optional()
        .describe(
          "Folder name: inbox, sent, drafts, deleted, junk, archive, or a folder ID. Defaults to inbox.",
        ),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Number of messages to return (1–100). Defaults to 10."),
      include_body: z
        .boolean()
        .optional()
        .describe("Include message body. Defaults to true."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const folder = resolveFolder(params.folder ?? "inbox");
      const limit = params.limit ?? 10;
      const select = [
        "id",
        "subject",
        "from",
        "toRecipients",
        "receivedDateTime",
        "isRead",
        "hasAttachments",
        "conversationId",
        "bodyPreview",
      ];
      if (params.include_body !== false) select.push("body");

      const data = await graph.request(
        "GET",
        `/me/mailFolders/${folder}/messages`,
        undefined,
        {
          $top: String(limit),
          $select: select.join(","),
          $orderby: "receivedDateTime desc",
        },
      );
      return data.value ?? [];
    },
  }),

  defineTool({
    name: "get_email",
    description: "Get a single email message by ID with full body and headers.",
    inputSchema: {
      email_id: z.string().describe("Message ID"),
      include_body: z
        .boolean()
        .optional()
        .describe("Include message body. Defaults to true."),
      body_max_length: z
        .number()
        .int()
        .optional()
        .describe("Max characters of body content to return. Defaults to 50000."),
      include_attachments: z
        .boolean()
        .optional()
        .describe("Include attachment metadata. Defaults to true."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph, log }, params) => {
      // Enhanced validation to catch malformed IDs early
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        // Log the invalid ID attempt for debugging
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      const select = [
        "id",
        "subject",
        "from",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "receivedDateTime",
        "sentDateTime",
        "isRead",
        "hasAttachments",
        "conversationId",
        "internetMessageId",
      ];
      if (params.include_body !== false) select.push("body");

      const queryParams: Record<string, string> = { $select: select.join(",") };
      if (params.include_attachments !== false) {
        queryParams.$expand =
          "attachments($select=id,name,contentType,size,isInline)";
      }

      const data = await emailReq(
        graph.request("GET", `/me/messages/${params.email_id}`, undefined, queryParams),
      );

      const maxLen = params.body_max_length ?? 50000;
      if (data.body?.content) data.body.content = data.body.content.slice(0, maxLen);

      return data;
    },
  }),

  defineTool({
    name: "create_email_draft",
    description:
      "Create a new, standalone draft email that can be edited and sent later. Use this only for starting a fresh message — to draft a reply within an existing conversation thread, use create_reply_draft instead.",
    inputSchema: {
      to: z.array(recipient).describe("Recipient email addresses"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body"),
      body_type: z
        .enum(["text", "html"])
        .optional()
        .describe("Body content type. Defaults to text."),
      cc: z.array(recipient).optional().describe("CC recipients"),
      bcc: z.array(recipient).optional().describe("BCC recipients"),
      attachments: z.array(attachment).optional().describe("File attachments"),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) =>
      graph.request("POST", "/me/messages", buildEmailBody(params)),
  }),

  defineTool({
    name: "create_reply_draft",
    description:
      "Draft a reply to an existing email, kept inside the original conversation thread. Produces an editable draft (in Drafts) that inherits the conversation, the 'RE:' subject, threading headers, and quoted history — review and send it later with send_email or via Outlook. Use this instead of create_email_draft whenever responding to a message.",
    inputSchema: {
      email_id: z.string().describe("Message ID being replied to"),
      body: z
        .string()
        .describe("Your reply text, inserted above the quoted original message"),
      reply_all: z
        .boolean()
        .optional()
        .describe(
          "Reply to all recipients of the original (sender + to + cc) instead of just the sender. Defaults to false.",
        ),
      cc: z
        .array(recipient)
        .optional()
        .describe("Additional CC recipients to add to the reply"),
    },
    confirmationPolicy: "always",
    handler: async ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }

      const endpoint = params.reply_all ? "createReplyAll" : "createReply";
      const body: any = { comment: params.body };
      if (params.cc?.length) body.message = { ccRecipients: toRecipients(params.cc) };

      const draft = await emailReq(
        graph.request("POST", `/me/messages/${params.email_id}/${endpoint}`, body),
      );

      // Return a concise summary instead of the full draft, whose body contains
      // the entire quoted conversation history.
      return {
        id: draft.id,
        conversationId: draft.conversationId,
        subject: draft.subject,
        webLink: draft.webLink,
        isDraft: draft.isDraft,
      };
    },
  }),

  defineTool({
    name: "send_email",
    description: "Send an email immediately.",
    inputSchema: {
      to: z.array(recipient).describe("Recipient email addresses"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body"),
      body_type: z
        .enum(["text", "html"])
        .optional()
        .describe("Body content type. Defaults to text."),
      cc: z.array(recipient).optional().describe("CC recipients"),
      bcc: z.array(recipient).optional().describe("BCC recipients"),
      attachments: z.array(attachment).optional().describe("File attachments"),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      await graph.request("POST", "/me/sendMail", {
        message: buildEmailBody(params),
        saveToSentItems: true,
      });
      return { success: true };
    },
  }),

  defineTool({
    name: "update_email",
    description:
      "Update email properties such as read status, categories, or flag.",
    inputSchema: {
      email_id: z.string().describe("Message ID to update"),
      updates: z
        .record(z.string(), z.unknown())
        .describe(
          'Properties to update, e.g. {"isRead": true, "categories": ["work"]}',
        ),
    },
    confirmationPolicy: "always",
    handler: ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      return emailReq(graph.request("PATCH", `/me/messages/${params.email_id}`, params.updates));
    },
  }),

  defineTool({
    name: "delete_email",
    description: "Permanently delete an email message.",
    inputSchema: {
      email_id: z.string().describe("Message ID to delete"),
    },
    confirmationPolicy: "always",
    handler: async ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      await emailReq(graph.request("DELETE", `/me/messages/${params.email_id}`));
      return { success: true };
    },
  }),

  defineTool({
    name: "move_email",
    description: "Move an email to a different folder.",
    inputSchema: {
      email_id: z.string().describe("Message ID to move"),
      destination_folder: z
        .string()
        .describe(
          "Destination folder name (inbox, sent, drafts, deleted, junk, archive) or folder ID",
        ),
    },
    confirmationPolicy: "always",
    handler: ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      return emailReq(
        graph.request("POST", `/me/messages/${params.email_id}/move`, {
          destinationId: resolveFolder(params.destination_folder),
        }),
      );
    },
  }),

  defineTool({
    name: "reply_to_email",
    description: "Reply to an email (sender only).",
    inputSchema: {
      email_id: z.string().describe("Message ID to reply to"),
      body: z.string().describe("Reply body"),
      body_type: z
        .enum(["text", "html"])
        .optional()
        .describe("Body content type. Defaults to text."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      await emailReq(
        graph.request("POST", `/me/messages/${params.email_id}/reply`, {
          message: {
            body: {
              contentType: params.body_type === "html" ? "HTML" : "Text",
              content: params.body,
            },
          },
        }),
      );
      return { success: true };
    },
  }),

  defineTool({
    name: "reply_all_email",
    description: "Reply to all recipients of an email.",
    inputSchema: {
      email_id: z.string().describe("Message ID to reply-all to"),
      body: z.string().describe("Reply body"),
      body_type: z
        .enum(["text", "html"])
        .optional()
        .describe("Body content type. Defaults to text."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
      } catch (error: any) {
        log.warn("Invalid email_id format", { email_id: params.email_id, error: error.message });
        throw error;
      }
      await emailReq(
        graph.request("POST", `/me/messages/${params.email_id}/replyAll`, {
          message: {
            body: {
              contentType: params.body_type === "html" ? "HTML" : "Text",
              content: params.body,
            },
          },
        }),
      );
      return { success: true };
    },
  }),

  defineTool({
    name: "get_attachment",
    description:
      "Get an email attachment's content (returned as base64 in contentBytes).",
    inputSchema: {
      email_id: z.string().describe("Message ID"),
      attachment_id: z.string().describe("Attachment ID"),
    },
    confirmationPolicy: "never",
    handler: ({ graph, log }, params) => {
      try {
        validateId(params.email_id, "email_id");
        validateId(params.attachment_id, "attachment_id");
      } catch (error: any) {
        log.warn("Invalid ID format", {
          email_id: params.email_id,
          attachment_id: params.attachment_id,
          error: error.message,
        });
        throw error;
      }
      return emailReq(
        graph.request(
          "GET",
          `/me/messages/${params.email_id}/attachments/${params.attachment_id}`,
        ),
      );
    },
  }),

  defineTool({
    name: "search_emails",
    description:
      "Search emails by keyword across all folders or within a specific folder (KQL syntax supported).",
    inputSchema: {
      query: z.string().describe("Search query string"),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max results. Defaults to 50."),
      folder: z
        .string()
        .optional()
        .describe("Limit search to this folder (inbox, sent, etc.)"),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const limit = params.limit ?? 50;

      if (params.folder) {
        const folder = resolveFolder(params.folder);
        const data = await graph.request(
          "GET",
          `/me/mailFolders/${folder}/messages`,
          undefined,
          {
            $search: `"${escapeKql(params.query)}"`,
            $top: String(limit),
            $select:
              "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview",
          },
        );
        return data.value ?? [];
      }

      const result = await graph.request("POST", "/search/query", {
        requests: [
          {
            entityTypes: ["message"],
            query: { queryString: params.query },
            from: 0,
            size: Math.min(limit, 25),
            fields: [
              "id",
              "subject",
              "from",
              "toRecipients",
              "receivedDateTime",
              "isRead",
              "bodyPreview",
            ],
          },
        ],
      });

      return (
        result?.value?.[0]?.hitsContainers?.[0]?.hits?.map(
          (h: any) => h.resource,
        ) ?? []
      );
    },
  }),

  defineTool({
    name: "unified_search",
    description:
      "Search across multiple Microsoft 365 content types (email, calendar events, files) in one query.",
    inputSchema: {
      query: z.string().describe("Search query string"),
      entity_types: z
        .array(z.enum(["message", "event", "driveItem", "site", "listItem"]))
        .optional()
        .describe(
          "Entity types to search. Defaults to [message, event, driveItem].",
        ),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max total results. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const entityTypes = params.entity_types ?? ["message", "event", "driveItem"];
      const limit = params.limit ?? 50;

      // Validate entity type combinations
      // Microsoft Graph has restrictions on which entity types can be searched together
      const validCombinations = [
        ["message"],
        ["event"],
        ["driveItem"],
        ["message", "event"],
        ["driveItem", "site"],
        ["driveItem", "listItem"],
        ["site", "listItem"],
        ["driveItem", "site", "listItem"],
      ];

      const isValidCombination = validCombinations.some(
        (combo) =>
          entityTypes.length === combo.length &&
          entityTypes.every((type: string) => combo.includes(type)),
      );

      if (!isValidCombination) {
        throw httpError(
          `Invalid entity type combination. Valid combinations: ${validCombinations.map((c) => c.join("+")).join(", ")}`,
          400,
        );
      }

      const result = await graph.request("POST", "/search/query", {
        requests: [
          {
            entityTypes,
            query: { queryString: params.query },
            from: 0,
            size: Math.min(limit, 25),
          },
        ],
      });

      const hits: any[] = [];
      for (const container of result?.value ?? []) {
        for (const hitsContainer of container?.hitsContainers ?? []) {
          for (const hit of hitsContainer?.hits ?? []) {
            hits.push({
              type: hit["@odata.type"]?.replace("#microsoft.graph.", "") ?? "unknown",
              score: hit._score,
              ...hit.resource,
            });
          }
        }
      }

      return hits;
    },
  }),
];
