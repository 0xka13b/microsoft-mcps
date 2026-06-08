import { describe, it, expect } from "vitest";
import { tools } from "./tools.js";

const EXPECTED_TOOL_COUNT = 9;
const NAME_PATTERN = /^[a-z][a-z0-9_]*$/;

describe("onedrive tools", () => {
  it("is a non-empty array", () => {
    expect(Array.isArray(tools)).toBe(true);
    expect(tools.length).toBeGreaterThan(0);
  });

  it("locks the exact public tool count", () => {
    expect(tools.length).toBe(EXPECTED_TOOL_COUNT);
  });

  it("has unique tool names", () => {
    const names = tools.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });

  for (const tool of tools) {
    describe(`tool: ${tool.name}`, () => {
      it("has a valid snake_case name", () => {
        expect(typeof tool.name).toBe("string");
        expect(tool.name.length).toBeGreaterThan(0);
        expect(tool.name).toMatch(NAME_PATTERN);
      });

      it("has a non-empty description", () => {
        expect(typeof tool.description).toBe("string");
        expect(tool.description.length).toBeGreaterThan(0);
      });

      it("has a valid confirmationPolicy", () => {
        expect(["always", "never"]).toContain(tool.confirmationPolicy);
      });

      it("has a handler function", () => {
        expect(typeof tool.handler).toBe("function");
      });

      it("has a non-null object inputSchema", () => {
        expect(tool.inputSchema).not.toBeNull();
        expect(typeof tool.inputSchema).toBe("object");
      });
    });
  }
});
