import { z } from "zod";
import type { GraphClient } from "@microsoft-mcp/graph";
import type { Logger } from "@microsoft-mcp/logger";

/** Identifying metadata for an MCP server, surfaced to clients on connect. */
export interface ServerMeta {
  /** Stable machine name, e.g. "microsoft-calendar". */
  name: string;
  version: string;
  /** Optional human-friendly display name. */
  title?: string;
}

/** Per-invocation context handed to every tool handler. */
export interface ToolContext {
  /** Microsoft Graph client pre-bound to the caller's access token. */
  graph: GraphClient;
  log: Logger;
}

/**
 * Whether a client should confirm with the user before this tool runs.
 * `always` marks mutating/destructive tools; `never` marks read-only tools.
 * Mapped onto MCP tool annotations (readOnlyHint / destructiveHint).
 */
export type ConfirmationPolicy = "always" | "never";

/**
 * Declarative definition of a single tool: its schema, metadata, and handler
 * co-located. Register a collection of these with {@link run}.
 */
export interface ToolDefinition<Shape extends z.ZodRawShape = z.ZodRawShape> {
  /** Tool name exposed to clients (snake_case), e.g. "list_events". */
  name: string;
  description: string;
  /** Raw Zod shape; the SDK validates incoming arguments against it. */
  inputSchema: Shape;
  confirmationPolicy: ConfirmationPolicy;
  /** Business logic. Receives validated, typed arguments. */
  handler: (ctx: ToolContext, args: z.infer<z.ZodObject<Shape>>) => Promise<unknown>;
}

/** A tool definition with its schema generic erased — used for storage/registration. */
export type AnyTool = ToolDefinition<any>;

/**
 * Identity helper that preserves the Zod shape so handler `args` are fully
 * typed at the definition site. Collect the results into a `tools` array.
 */
export function defineTool<Shape extends z.ZodRawShape>(
  def: ToolDefinition<Shape>,
): ToolDefinition<Shape> {
  return def;
}
