import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { log } from "./index.js";

describe("logger", () => {
  let writeSpy: ReturnType<typeof vi.spyOn>;
  let stdoutSpy: ReturnType<typeof vi.spyOn>;

  const originalDebug = process.env.MCP_DEBUG;

  beforeEach(() => {
    writeSpy = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    stdoutSpy = vi.spyOn(process.stdout, "write").mockReturnValue(true);
  });

  afterEach(() => {
    writeSpy.mockRestore();
    stdoutSpy.mockRestore();
    vi.restoreAllMocks();
    if (originalDebug === undefined) {
      delete process.env.MCP_DEBUG;
    } else {
      process.env.MCP_DEBUG = originalDebug;
    }
  });

  const parseLine = (raw: unknown) => {
    const line = raw as string;
    expect(typeof line).toBe("string");
    expect(line.endsWith("\n")).toBe(true);
    return JSON.parse(line.trimEnd()) as Record<string, unknown>;
  };

  describe.each([
    ["info", () => log.info("hello info")] as const,
    ["warn", () => log.warn("hello warn")] as const,
    ["error", () => log.error("hello error")] as const,
  ])("log.%s", (level, call) => {
    it(`writes exactly one line to stderr with level "${level}"`, () => {
      call();

      expect(writeSpy).toHaveBeenCalledTimes(1);

      const parsed = parseLine(writeSpy.mock.calls[0][0]);
      expect(parsed.level).toBe(level);
      expect(parsed.msg).toBe(`hello ${level}`);
      expect(typeof parsed.ts).toBe("string");
      expect(Number.isNaN(Date.parse(parsed.ts as string))).toBe(false);
    });

    it(`does not write to stdout for level "${level}"`, () => {
      call();
      expect(stdoutSpy).toHaveBeenCalledTimes(0);
    });
  });

  it("merges extra data fields into the JSON for log.info", () => {
    log.info("hi", { tool: "x", n: 1 });

    expect(writeSpy).toHaveBeenCalledTimes(1);

    const parsed = parseLine(writeSpy.mock.calls[0][0]);
    expect(parsed.level).toBe("info");
    expect(parsed.msg).toBe("hi");
    expect(parsed.tool).toBe("x");
    expect(parsed.n).toBe(1);
    expect(Number.isNaN(Date.parse(parsed.ts as string))).toBe(false);
    expect(stdoutSpy).toHaveBeenCalledTimes(0);
  });

  it("merges extra data fields for log.warn and log.error", () => {
    log.warn("warn msg", { code: 42 });
    log.error("error msg", { reason: "boom" });

    expect(writeSpy).toHaveBeenCalledTimes(2);

    const warnParsed = parseLine(writeSpy.mock.calls[0][0]);
    expect(warnParsed.level).toBe("warn");
    expect(warnParsed.msg).toBe("warn msg");
    expect(warnParsed.code).toBe(42);

    const errorParsed = parseLine(writeSpy.mock.calls[1][0]);
    expect(errorParsed.level).toBe("error");
    expect(errorParsed.msg).toBe("error msg");
    expect(errorParsed.reason).toBe("boom");

    expect(stdoutSpy).toHaveBeenCalledTimes(0);
  });

  describe("log.debug", () => {
    it("writes nothing when MCP_DEBUG is unset", () => {
      delete process.env.MCP_DEBUG;

      log.debug("should not appear", { hidden: true });

      expect(writeSpy).toHaveBeenCalledTimes(0);
      expect(stdoutSpy).toHaveBeenCalledTimes(0);
    });

    it("writes nothing when MCP_DEBUG is an empty string (falsy)", () => {
      process.env.MCP_DEBUG = "";

      log.debug("should not appear");

      expect(writeSpy).toHaveBeenCalledTimes(0);
    });

    it('writes a debug line when MCP_DEBUG is set to "1"', () => {
      process.env.MCP_DEBUG = "1";

      log.debug("debug on", { detail: "yes" });

      expect(writeSpy).toHaveBeenCalledTimes(1);

      const parsed = parseLine(writeSpy.mock.calls[0][0]);
      expect(parsed.level).toBe("debug");
      expect(parsed.msg).toBe("debug on");
      expect(parsed.detail).toBe("yes");
      expect(Number.isNaN(Date.parse(parsed.ts as string))).toBe(false);
      expect(stdoutSpy).toHaveBeenCalledTimes(0);
    });
  });
});
