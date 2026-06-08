import { describe, it, expect } from "vitest";
import { homedir } from "node:os";
import { join } from "node:path";
import {
  resolveAuthConfig,
  authorityFor,
  configDir,
  cacheFilePath,
  AuthConfigError,
  NotSignedInError,
  CLIENT_ID_ENV,
  TENANT_ENV,
  CACHE_DIR_ENV,
} from "./index.js";

const AUTH = "https://login.microsoftonline.com";

describe("authorityFor", () => {
  it("defaults to common when undefined", () => {
    expect(authorityFor(undefined)).toBe(`${AUTH}/common`);
  });
  it("uses a given tenant", () => {
    expect(authorityFor("contoso.onmicrosoft.com")).toBe(`${AUTH}/contoso.onmicrosoft.com`);
  });
  it("treats empty / whitespace as common", () => {
    expect(authorityFor("")).toBe(`${AUTH}/common`);
    expect(authorityFor("   ")).toBe(`${AUTH}/common`);
  });
  it("trims surrounding whitespace", () => {
    expect(authorityFor("  organizations  ")).toBe(`${AUTH}/organizations`);
  });
});

describe("configDir / cacheFilePath", () => {
  it("uses CACHE_DIR_ENV when set", () => {
    expect(configDir({ [CACHE_DIR_ENV]: "/tmp/cache" })).toBe("/tmp/cache");
    expect(cacheFilePath({ [CACHE_DIR_ENV]: "/tmp/cache" })).toBe(join("/tmp/cache", "token-cache.json"));
  });
  it("uses XDG_CONFIG_HOME next", () => {
    expect(configDir({ XDG_CONFIG_HOME: "/home/u/.cfg" })).toBe(join("/home/u/.cfg", "microsoft-mcp"));
  });
  it("falls back to ~/.config/microsoft-mcp", () => {
    expect(configDir({})).toBe(join(homedir(), ".config", "microsoft-mcp"));
  });
  it("CACHE_DIR_ENV wins over XDG_CONFIG_HOME", () => {
    expect(configDir({ [CACHE_DIR_ENV]: "/a", XDG_CONFIG_HOME: "/b" })).toBe("/a");
  });
});

describe("resolveAuthConfig", () => {
  it("throws AuthConfigError (401) mentioning the env var when client id is missing", () => {
    let caught: any;
    try {
      resolveAuthConfig({});
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(AuthConfigError);
    expect(caught.status).toBe(401);
    expect(caught.message).toContain(CLIENT_ID_ENV);
  });

  it("returns config with default authority + cache file", () => {
    const c = resolveAuthConfig({ [CLIENT_ID_ENV]: "abc-123" });
    expect(c.clientId).toBe("abc-123");
    expect(c.authority).toBe(`${AUTH}/common`);
    expect(c.cacheFile).toBe(join(homedir(), ".config", "microsoft-mcp", "token-cache.json"));
  });

  it("honors tenant + cache-dir overrides", () => {
    const c = resolveAuthConfig({
      [CLIENT_ID_ENV]: "abc-123",
      [TENANT_ENV]: "my-tenant-id",
      [CACHE_DIR_ENV]: "/tmp/c",
    });
    expect(c.authority).toBe(`${AUTH}/my-tenant-id`);
    expect(c.cacheFile).toBe(join("/tmp/c", "token-cache.json"));
  });
});

describe("error types", () => {
  it("AuthConfigError is a 401 Error", () => {
    const e = new AuthConfigError("nope");
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe("AuthConfigError");
    expect(e.status).toBe(401);
  });
  it("NotSignedInError is a 401 Error", () => {
    const e = new NotSignedInError("nope");
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe("NotSignedInError");
    expect(e.status).toBe(401);
  });
});
