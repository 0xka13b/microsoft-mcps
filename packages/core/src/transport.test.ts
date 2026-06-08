import { describe, it, expect } from "vitest";
import { extractBearer, requireToken, TOKEN_ENV } from "./transport.js";

describe("extractBearer", () => {
  it("extracts token from 'Bearer abc123'", () => {
    expect(extractBearer("Bearer abc123")).toBe("abc123");
  });

  it("trims surrounding whitespace", () => {
    expect(extractBearer("Bearer   abc  ")).toBe("abc");
  });

  it("returns undefined when empty after trim", () => {
    expect(extractBearer("Bearer ")).toBeUndefined();
  });

  it("returns undefined for non-Bearer scheme", () => {
    expect(extractBearer("Basic xyz")).toBeUndefined();
  });

  it("returns undefined for undefined header", () => {
    expect(extractBearer(undefined)).toBeUndefined();
  });

  it("returns undefined for lowercase 'bearer' (must start with 'Bearer ')", () => {
    expect(extractBearer("bearer abc")).toBeUndefined();
  });
});

describe("requireToken", () => {
  it("returns the token for a non-empty string", () => {
    expect(requireToken("tok")).toBe("tok");
  });

  it("throws a 401 error mentioning TOKEN_ENV when undefined", () => {
    expect(TOKEN_ENV).toBe("MICROSOFT_ACCESS_TOKEN");
    let thrown: any;
    try {
      requireToken(undefined);
    } catch (err) {
      thrown = err;
    }
    expect(thrown).toBeInstanceOf(Error);
    expect(thrown.status).toBe(401);
    expect(thrown.message).toContain(TOKEN_ENV);
  });
});
