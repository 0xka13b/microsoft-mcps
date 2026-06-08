import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { resolveMode, resolvePort } from "./run.js";

describe("resolveMode", () => {
  let savedTransport: string | undefined;

  beforeEach(() => {
    savedTransport = process.env.MCP_TRANSPORT;
    delete process.env.MCP_TRANSPORT;
  });

  afterEach(() => {
    if (savedTransport === undefined) delete process.env.MCP_TRANSPORT;
    else process.env.MCP_TRANSPORT = savedTransport;
  });

  it("returns 'http' for --http flag", () => {
    expect(resolveMode(["--http"])).toBe("http");
  });

  it("returns 'stdio' for --stdio flag", () => {
    expect(resolveMode(["--stdio"])).toBe("stdio");
  });

  it("--http flag beats env MCP_TRANSPORT=stdio", () => {
    process.env.MCP_TRANSPORT = "stdio";
    expect(resolveMode(["--http"])).toBe("http");
  });

  it("uses env MCP_TRANSPORT=http when no flag", () => {
    process.env.MCP_TRANSPORT = "http";
    expect(resolveMode([])).toBe("http");
  });

  it("env MCP_TRANSPORT is case-insensitive (HTTP)", () => {
    process.env.MCP_TRANSPORT = "HTTP";
    expect(resolveMode([])).toBe("http");
  });

  it("returns 'stdio' for env MCP_TRANSPORT=stdio", () => {
    process.env.MCP_TRANSPORT = "stdio";
    expect(resolveMode([])).toBe("stdio");
  });

  it("returns 'stdio' when env unset and no flag", () => {
    expect(resolveMode([])).toBe("stdio");
  });

  it("returns 'stdio' for garbage env value", () => {
    process.env.MCP_TRANSPORT = "garbage";
    expect(resolveMode([])).toBe("stdio");
  });

  it("--http beats --stdio when both present (http checked first)", () => {
    expect(resolveMode(["--stdio", "--http"])).toBe("http");
  });
});

describe("resolvePort", () => {
  let savedPort: string | undefined;

  beforeEach(() => {
    savedPort = process.env.PORT;
    delete process.env.PORT;
  });

  afterEach(() => {
    if (savedPort === undefined) delete process.env.PORT;
    else process.env.PORT = savedPort;
  });

  it("parses --port=8080", () => {
    expect(resolvePort(["--port=8080"])).toBe(8080);
  });

  it("parses --port 8080 (separate args)", () => {
    expect(resolvePort(["--port", "8080"])).toBe(8080);
  });

  it("flag beats PORT env", () => {
    process.env.PORT = "5000";
    expect(resolvePort(["--port=8080"])).toBe(8080);
  });

  it("uses PORT env when no flag", () => {
    process.env.PORT = "5000";
    expect(resolvePort([])).toBe(5000);
  });

  it("defaults to 3000 when nothing set", () => {
    expect(resolvePort([])).toBe(3000);
  });
});
