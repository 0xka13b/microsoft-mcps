import { log } from "@microsoft-mcp/logger";
import { login, resolveAuthConfig } from "@microsoft-mcp/auth";
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

/* v8 ignore start -- interactive sign-in (browser / device code), not unit-testable */
/**
 * Runs the one-time interactive sign-in: opens the browser (or device-code flow
 * with `--device-code`) and caches the token so the server can run silently.
 */
async function runLogin(meta: ServerMeta, argv: string[]): Promise<void> {
  try {
    const config = resolveAuthConfig();
    const scopes = meta.scopes ?? [];
    console.log(`Signing in to Microsoft for ${meta.title ?? meta.name}...`);
    if (scopes.length) console.log(`Requesting scopes: ${scopes.join(", ")}`);
    await login(config, scopes, argv.includes("--device-code"));
    console.log(
      "\n✓ Signed in. Your token is cached and refreshed automatically — " +
        "start the server normally (no MICROSOFT_ACCESS_TOKEN needed).",
    );
  } catch (err: any) {
    console.error(`\nSign-in failed: ${err?.message ?? err}`);
    process.exitCode = 1;
  }
}
/* v8 ignore stop */

/**
 * Entry point for every Microsoft MCP server. Handles the `login` subcommand,
 * otherwise wires the tool collection to the selected transport.
 */
export async function run(meta: ServerMeta, tools: AnyTool[]): Promise<void> {
  const argv = process.argv.slice(2);

  if (argv[0] === "login") {
    await runLogin(meta, argv);
    return;
  }

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
