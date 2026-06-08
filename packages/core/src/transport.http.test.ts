import type { Server } from "node:http";
import { describe, it, expect, beforeAll, afterAll, vi } from "vitest";
import request from "supertest";
import { z } from "zod";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { defineTool } from "./types.js";
import { runHttp } from "./transport.js";

vi.spyOn(process.stderr, "write").mockReturnValue(true);

const echoTool = defineTool({
  name: "echo",
  description: "Echoes the query back to the caller.",
  confirmationPolicy: "never",
  inputSchema: { q: z.string() },
  handler: async (_ctx, args) => ({ ok: true, q: args.q }),
});

const tools = [echoTool];

let server: Server;
let port: number;

beforeAll(async () => {
  server = await runHttp({ name: "test-http", version: "1.0.0" }, tools, 0);
  port = (server.address() as { port: number }).port;
});

afterAll(async () => {
  await new Promise<void>((r) => server.close(() => r()));
});

describe("HTTP transport — supertest", () => {
  it("GET /healthz returns ok status with server metadata", async () => {
    const res = await request(server).get("/healthz");
    expect(res.status).toBe(200);
    expect(res.body.status).toBe("ok");
    expect(res.body.server).toBe("test-http");
    expect(res.body.version).toBe("1.0.0");
  });

  it("GET /mcp returns 405 with a JSON-RPC method-not-allowed error", async () => {
    const res = await request(server).get("/mcp");
    expect(res.status).toBe(405);
    expect(res.body.error.message).toContain("Method not allowed");
  });

  it("DELETE /mcp returns 405", async () => {
    const res = await request(server).delete("/mcp");
    expect(res.status).toBe(405);
  });
});

describe("HTTP transport — MCP round-trip over Streamable HTTP", () => {
  let client: Client;
  let transport: StreamableHTTPClientTransport;

  beforeAll(async () => {
    transport = new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`), {
      requestInit: { headers: { Authorization: "Bearer test-token" } },
    });
    client = new Client({ name: "test-client", version: "1.0.0" });
    await client.connect(transport);
  });

  afterAll(async () => {
    await client.close();
    await transport.close();
  });

  it("listTools exposes the read-only tool with readOnlyHint", async () => {
    const { tools: listed } = await client.listTools();
    const tool = listed.find((t) => t.name === "echo");
    expect(tool).toBeDefined();
    expect(tool?.annotations?.readOnlyHint).toBe(true);
  });

  it("callTool runs the handler and returns the echoed payload", async () => {
    const result = (await client.callTool({
      name: "echo",
      arguments: { q: "hello" },
    })) as { isError?: boolean; content: Array<{ type: string; text: string }> };

    expect(result.isError).toBeFalsy();
    expect(JSON.parse(result.content[0].text)).toEqual({ ok: true, q: "hello" });
  });
});
