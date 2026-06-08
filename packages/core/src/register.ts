import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { createGraphClient } from "@microsoft-mcp/graph";
import { log } from "@microsoft-mcp/logger";
import type { AnyTool, ServerMeta, ToolContext } from "./types.js";

/** Supplies the Microsoft Graph access token for a request, or throws if absent. */
export type TokenProvider = () => string;

const jsonResult = (value: unknown): CallToolResult => ({
  content: [{ type: "text", text: JSON.stringify(value ?? null, null, 2) }],
});

const errorResult = (message: string, status: number): CallToolResult => ({
  content: [{ type: "text", text: JSON.stringify({ error: message, status }, null, 2) }],
  isError: true,
});

/**
 * Maps an error thrown by a tool handler (or the Graph client) onto an MCP
 * error result with status-specific messaging (auth, rate limit, client vs.
 * server errors).
 *
 * Exported for unit testing.
 */
export const mapError = (toolName: string, err: any, durationMs: number): CallToolResult => {
  const status: number = err?.status ?? err?.code ?? 500;
  const message: string = err?.message ?? "Tool execution failed";

  if (status === 401 || status === 403) {
    log.warn("tool auth error", { tool: toolName, status, message, durationMs });
    return errorResult(`Authentication failed: ${message}`, status);
  }
  if (status === 429) {
    log.warn("tool rate limit", { tool: toolName, status, durationMs });
    return errorResult("Rate limit exceeded. Please retry after a short delay.", 429);
  }
  if (status >= 500 && status < 600) {
    log.error("tool server error", { tool: toolName, status, message, durationMs });
    return errorResult("Microsoft Graph service error. Please try again later.", status);
  }
  if (status >= 400 && status < 500) {
    log.error("tool client error", { tool: toolName, status, message, durationMs });
    return errorResult(message || "Invalid request", status);
  }

  log.error("tool error", { tool: toolName, status, message, durationMs });
  return errorResult(message, status);
};

/**
 * Builds an {@link McpServer} with every tool registered. The `getToken`
 * provider is invoked per call, so a single server definition serves both the
 * stdio transport (env-var token) and the HTTP transport (per-request token).
 */
export function createServer(meta: ServerMeta, tools: AnyTool[], getToken: TokenProvider): McpServer {
  const server = new McpServer({
    name: meta.name,
    version: meta.version,
    ...(meta.title ? { title: meta.title } : {}),
  });

  for (const tool of tools) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.inputSchema,
        annotations: {
          readOnlyHint: tool.confirmationPolicy === "never",
          destructiveHint: tool.confirmationPolicy === "always",
        },
      },
      async (args: any): Promise<CallToolResult> => {
        const start = Date.now();

        let token: string;
        try {
          token = getToken();
        } catch (err: any) {
          log.warn("tool auth missing", { tool: tool.name, message: err?.message });
          return errorResult(err?.message ?? "Authentication required", err?.status ?? 401);
        }

        const ctx: ToolContext = { graph: createGraphClient(token), log };
        log.info("tool call", { tool: tool.name });

        try {
          const result = await tool.handler(ctx, args);
          log.info("tool success", { tool: tool.name, durationMs: Date.now() - start });
          return jsonResult(result);
        } catch (err: any) {
          return mapError(tool.name, err, Date.now() - start);
        }
      },
    );
  }

  return server;
}
