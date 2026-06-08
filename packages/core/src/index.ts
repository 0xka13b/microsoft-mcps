export { run } from "./run.js";
export type { TransportMode } from "./run.js";
export { createServer } from "./register.js";
export type { TokenProvider } from "./register.js";
export { runStdio, runHttp, TOKEN_ENV } from "./transport.js";
export { defineTool } from "./types.js";
export type {
  AnyTool,
  ConfirmationPolicy,
  ServerMeta,
  ToolContext,
  ToolDefinition,
} from "./types.js";

// Re-export the Graph client surface so tool authors can import everything from
// "@microsoft-mcp/core" without a separate dependency line when they prefer.
export { GraphClient, GraphError, createGraphClient, GRAPH_BASE } from "@microsoft-mcp/graph";
export type { QueryParams, RawRequestOptions } from "@microsoft-mcp/graph";
