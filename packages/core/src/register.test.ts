import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { z } from "zod";
import { createServer, mapError } from "./register.js";
import { defineTool } from "./types.js";

const parse = (result: any): { error: string; status: number } =>
  JSON.parse(result.content[0].text);

describe("mapError", () => {
  let stderrSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    stderrSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);
  });

  afterEach(() => {
    stderrSpy.mockRestore();
  });

  it("maps 401 to an authentication failure", () => {
    const result = mapError("t", { status: 401, message: "bad" }, 1);
    expect(result.isError).toBe(true);
    const parsed = parse(result);
    expect(parsed.status).toBe(401);
    expect(parsed.error.startsWith("Authentication failed:")).toBe(true);
  });

  it("maps 403 via the authentication path", () => {
    const result = mapError("t", { status: 403, message: "nope" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(403);
    expect(parsed.error.startsWith("Authentication failed:")).toBe(true);
  });

  it("maps 429 to the rate-limit message", () => {
    const result = mapError("t", { status: 429, message: "slow down" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(429);
    expect(parsed.error).toBe("Rate limit exceeded. Please retry after a short delay.");
  });

  it("maps 500 to a service-error message", () => {
    const result = mapError("t", { status: 500, message: "kaboom" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(500);
    expect(parsed.error).toBe("Microsoft Graph service error. Please try again later.");
  });

  it("maps 503 to a service-error message", () => {
    const result = mapError("t", { status: 503, message: "down" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(503);
    expect(parsed.error).toBe("Microsoft Graph service error. Please try again later.");
  });

  it("maps 400 (4xx) to the raw err.message", () => {
    const result = mapError("t", { status: 400, message: "invalid input" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(400);
    expect(parsed.error).toBe("invalid input");
  });

  it("maps 404 (4xx) to the raw err.message", () => {
    const result = mapError("t", { status: 404, message: "not found" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(404);
    expect(parsed.error).toBe("not found");
  });

  it("uses err.code as status when status is absent (4xx branch)", () => {
    const result = mapError("t", { code: 404, message: "missing" }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(404);
    expect(parsed.error).toBe("missing");
  });

  it("defaults to status 500 (service-error path) when neither status nor code present", () => {
    const result = mapError("t", {}, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(500);
    expect(parsed.error).toBe("Microsoft Graph service error. Please try again later.");
  });

  it("uses the 'Tool execution failed' fallback message for an out-of-range status", () => {
    // A status outside the 4xx/5xx branches falls through to the final
    // errorResult, surfacing the default message when err.message is absent.
    const result = mapError("t", { status: 200 }, 1);
    const parsed = parse(result);
    expect(parsed.status).toBe(200);
    expect(parsed.error).toBe("Tool execution failed");
  });
});

const tools = [
  defineTool({
    name: "read_tool",
    description: "A read-only tool",
    confirmationPolicy: "never",
    inputSchema: { q: z.string() },
    handler: async (_ctx, args) => ({ ok: true, q: args.q }),
  }),
  defineTool({
    name: "destructive_tool",
    description: "A destructive tool",
    confirmationPolicy: "always",
    inputSchema: {},
    handler: async () => {
      throw Object.assign(new Error("boom"), { status: 404 });
    },
  }),
];

describe("createServer integration", () => {
  let stderrSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    stderrSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);
  });

  afterEach(() => {
    stderrSpy.mockRestore();
  });

  async function connect(getToken: () => string) {
    const server = createServer({ name: "test", version: "1.0.0" }, tools, getToken);
    const [clientT, serverT] = InMemoryTransport.createLinkedPair();
    const client = new Client({ name: "c", version: "1.0.0" });
    await Promise.all([server.connect(serverT), client.connect(clientT)]);
    return { server, client };
  }

  it("exposes correct tool annotations", async () => {
    const { server, client } = await connect(() => "tok");
    const { tools: listed } = await client.listTools();

    const readTool = listed.find((t) => t.name === "read_tool");
    const destructiveTool = listed.find((t) => t.name === "destructive_tool");

    expect(readTool?.annotations?.readOnlyHint).toBe(true);
    expect(readTool?.annotations?.destructiveHint).toBe(false);
    expect(destructiveTool?.annotations?.readOnlyHint).toBe(false);
    expect(destructiveTool?.annotations?.destructiveHint).toBe(true);

    await client.close();
    await server.close();
  });

  it("runs a read-only tool and returns its result", async () => {
    const { server, client } = await connect(() => "tok");
    const result: any = await client.callTool({
      name: "read_tool",
      arguments: { q: "hi" },
    });

    expect(result.isError).toBeFalsy();
    expect(JSON.parse(result.content[0].text)).toEqual({ ok: true, q: "hi" });

    await client.close();
    await server.close();
  });

  it("maps a throwing tool to an error result", async () => {
    const { server, client } = await connect(() => "tok");
    const result: any = await client.callTool({
      name: "destructive_tool",
      arguments: {},
    });

    expect(result.isError).toBe(true);
    const parsed = parse(result);
    expect(parsed.status).toBe(404);
    expect(parsed.error).toBe("boom");

    await client.close();
    await server.close();
  });

  it("returns a 401 error when the token provider throws", async () => {
    const { server, client } = await connect(() => {
      throw Object.assign(new Error("no token"), { status: 401 });
    });

    const result: any = await client.callTool({
      name: "read_tool",
      arguments: { q: "hi" },
    });

    expect(result.isError).toBe(true);
    expect(parse(result).status).toBe(401);

    await client.close();
    await server.close();
  });
});
