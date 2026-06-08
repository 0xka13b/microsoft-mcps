import type { Server } from "node:http";
import express from "express";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { log } from "@microsoft-mcp/logger";
import { acquireTokenSilent, resolveAuthConfig } from "@microsoft-mcp/auth";
import { createServer, type TokenProvider } from "./register.js";
import type { AnyTool, ServerMeta } from "./types.js";

/** Env var holding the access token when running over stdio. */
export const TOKEN_ENV = "MICROSOFT_ACCESS_TOKEN";

export const requireToken = (token: string | undefined): string => {
  if (!token) {
    throw Object.assign(
      new Error(
        `Missing Microsoft Graph access token (set ${TOKEN_ENV} for stdio, or send 'Authorization: Bearer <token>' for HTTP).`,
      ),
      { status: 401 },
    );
  }
  return token;
};

export const extractBearer = (authHeader: string | undefined): string | undefined => {
  if (!authHeader?.startsWith("Bearer ")) return undefined;
  const token = authHeader.slice("Bearer ".length).trim();
  return token || undefined;
};

/**
 * Runs the server over stdio. Each tool call resolves a token: the
 * `MICROSOFT_ACCESS_TOKEN` env var wins (advanced/testing); otherwise an access
 * token is acquired silently from the cached MSAL sign-in (run the `login`
 * command once). The provider is per-call so refreshes happen transparently.
 */
export async function runStdio(meta: ServerMeta, tools: AnyTool[]): Promise<void> {
  const getToken: TokenProvider = async () => {
    const envToken = process.env[TOKEN_ENV];
    if (envToken) return envToken;
    return acquireTokenSilent(resolveAuthConfig(), meta.scopes ?? []);
  };

  const server = createServer(meta, tools, getToken);
  await server.connect(new StdioServerTransport());
  log.info("MCP server started", { server: meta.name, transport: "stdio" });
}

/**
 * Runs the server over Streamable HTTP (stateless). Each POST /mcp gets a fresh
 * server + transport with the request's bearer token, so distinct callers'
 * tokens never mix. GET/DELETE are unsupported in stateless mode.
 *
 * Resolves with the listening {@link Server} so callers (and tests) can close it.
 */
export async function runHttp(meta: ServerMeta, tools: AnyTool[], port: number): Promise<Server> {
  const app = express();
  app.use(express.json({ limit: process.env.MCP_HTTP_BODY_LIMIT ?? "50mb" }));

  app.get("/healthz", (_req, res) => {
    res.json({ status: "ok", server: meta.name, version: meta.version });
  });

  app.post("/mcp", async (req, res) => {
    const token = extractBearer(req.header("authorization"));
    const getToken: TokenProvider = () => requireToken(token);

    const server = createServer(meta, tools, getToken);
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });

    res.on("close", () => {
      void transport.close();
      void server.close();
    });

    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err: any) {
      log.error("http request error", { server: meta.name, message: err?.message });
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  const methodNotAllowed = (_req: express.Request, res: express.Response) => {
    res.status(405).json({
      jsonrpc: "2.0",
      error: { code: -32000, message: "Method not allowed (stateless server)." },
      id: null,
    });
  };
  app.get("/mcp", methodNotAllowed);
  app.delete("/mcp", methodNotAllowed);

  return await new Promise<Server>((resolve) => {
    const server = app.listen(port, () => {
      log.info("MCP server started", { server: meta.name, transport: "http", port });
      resolve(server);
    });
  });
}
