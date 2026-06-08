import { log } from "@microsoft-mcp/logger";
import { runHttp, runStdio } from "./transport.js";
import type { AnyTool, ServerMeta } from "./types.js";

export type TransportMode = "stdio" | "http";

/**
 * Selects the transport. Precedence:
 *   1. CLI flag `--stdio` / `--http`
 *   2. env `MCP_TRANSPORT=stdio|http`
 *   3. default `stdio` (the convention for MCP clients launching a subprocess)
 */
export const resolveMode = (argv: string[]): TransportMode => {
  if (argv.includes("--http")) return "http";
  if (argv.includes("--stdio")) return "stdio";
  const env = (process.env.MCP_TRANSPORT ?? "").toLowerCase();
  if (env === "http") return "http";
  return "stdio";
};

/**
 * Resolves the HTTP port. Precedence: `--port <n>` / `--port=<n>`, then `PORT`,
 * then 3000.
 */
export const resolvePort = (argv: string[]): number => {
  const eq = argv.find((a) => a.startsWith("--port="));
  if (eq) return Number(eq.slice("--port=".length));
  const idx = argv.indexOf("--port");
  if (idx >= 0 && argv[idx + 1]) return Number(argv[idx + 1]);
  if (process.env.PORT) return Number(process.env.PORT);
  return 3000;
};

/**
 * Entry point for every Microsoft MCP server. Wires the tool collection to the
 * selected transport and keeps the process alive.
 */
export async function run(meta: ServerMeta, tools: AnyTool[]): Promise<void> {
  const argv = process.argv.slice(2);
  const mode = resolveMode(argv);

  try {
    if (mode === "http") {
      await runHttp(meta, tools, resolvePort(argv));
    } else {
      await runStdio(meta, tools);
    }
  } catch (err: any) {
    log.error("fatal startup error", { server: meta.name, message: err?.message });
    process.exit(1);
  }
}
