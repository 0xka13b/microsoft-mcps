import { describe, it, expect } from "vitest";
import {
  httpError,
  validateId,
  validateDrivePath,
  escapeKql,
  escapeODataString,
} from "./index.js";

/**
 * Helper: captures a thrown error so its `.status` (and other props) can be
 * inspected. Asserts that the callback actually throws.
 */
const capture = (fn: () => unknown): Error & { status?: number } => {
  let thrown: unknown;
  let didThrow = false;
  try {
    fn();
  } catch (err) {
    didThrow = true;
    thrown = err;
  }
  expect(didThrow).toBe(true);
  return thrown as Error & { status?: number };
};

describe("httpError", () => {
  it("returns an Error instance", () => {
    const err = httpError("boom", 500);
    expect(err).toBeInstanceOf(Error);
  });

  it("sets the message", () => {
    const err = httpError("something failed", 400);
    expect(err.message).toBe("something failed");
  });

  it("attaches the status", () => {
    const err = httpError("nope", 401);
    expect(err.status).toBe(401);
  });

  it("supports an empty message", () => {
    const err = httpError("", 403);
    expect(err).toBeInstanceOf(Error);
    expect(err.message).toBe("");
    expect(err.status).toBe(403);
  });

  it("preserves arbitrary status codes", () => {
    const err = httpError("teapot", 418);
    expect(err.status).toBe(418);
  });
});

describe("validateId", () => {
  describe("happy path", () => {
    it("returns a realistic Graph id unchanged", () => {
      const id =
        "AAMkAGI1zABcDeFgHiJkLmNoPqRsTuVwXyZ0123456789-_aBcDeFgHiJkLmNoPqRsTuVwXyZ=";
      expect(validateId(id, "message id")).toBe(id);
    });

    it("returns a simple alphanumeric id unchanged", () => {
      expect(validateId("abc123", "id")).toBe("abc123");
    });

    it("allows ids containing dots", () => {
      expect(validateId("a.b.c", "id")).toBe("a.b.c");
    });

    it("allows ids containing a trailing equals (base64 padding)", () => {
      expect(validateId("AAMk==", "id")).toBe("AAMk==");
    });
  });

  describe("required / non-string", () => {
    it("throws on empty string", () => {
      const err = capture(() => validateId("", "message id"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("message id");
    });

    it("throws on a non-string value (null cast as any)", () => {
      const err = capture(() => validateId(null as any, "message id"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("message id");
    });

    it("throws on undefined cast as any", () => {
      const err = capture(() => validateId(undefined as any, "label"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("label");
    });

    it("throws on a number cast as any", () => {
      const err = capture(() => validateId(123 as any, "label"));
      expect(err.status).toBe(400);
    });

    it("also throws via expect().toThrow for empty string", () => {
      expect(() => validateId("", "id")).toThrow();
    });
  });

  describe("path-unsafe characters", () => {
    it("throws when containing a forward slash", () => {
      const err = capture(() => validateId("foo/bar", "id"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("id");
    });

    it("throws when containing a question mark", () => {
      const err = capture(() => validateId("foo?bar", "id"));
      expect(err.status).toBe(400);
    });

    it("throws when containing a hash", () => {
      const err = capture(() => validateId("foo#bar", "id"));
      expect(err.status).toBe(400);
    });

    it("throws when containing a NUL byte", () => {
      const err = capture(() => validateId("foo\x00bar", "id"));
      expect(err.status).toBe(400);
    });
  });

  describe("whitespace", () => {
    it("throws on a space", () => {
      const err = capture(() => validateId("foo bar", "id"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("id");
    });

    it("throws on a tab", () => {
      const err = capture(() => validateId("foo\tbar", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on a newline", () => {
      const err = capture(() => validateId("foo\nbar", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on a carriage return", () => {
      const err = capture(() => validateId("foo\rbar", "id"));
      expect(err.status).toBe(400);
    });
  });

  describe("malformed serialization values", () => {
    it("throws on 'undefined'", () => {
      const err = capture(() => validateId("undefined", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on 'UNDEFINED' (case-insensitive)", () => {
      const err = capture(() => validateId("UNDEFINED", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on 'null'", () => {
      const err = capture(() => validateId("null", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on 'Null' (case-insensitive)", () => {
      const err = capture(() => validateId("Null", "id"));
      expect(err.status).toBe(400);
    });

    it("throws on '[object Object]'", () => {
      const err = capture(() => validateId("[object Object]", "id"));
      expect(err.status).toBe(400);
    });
  });
});

describe("validateDrivePath", () => {
  describe("happy path", () => {
    it("returns a normal nested path unchanged", () => {
      expect(validateDrivePath("Documents/Reports/q1.txt", "path")).toBe(
        "Documents/Reports/q1.txt",
      );
    });

    it("allows a filename containing a dot (only whole segments rejected)", () => {
      expect(validateDrivePath("file.txt", "path")).toBe("file.txt");
    });

    it("allows a nested filename containing a dot", () => {
      expect(validateDrivePath("a/b/file.txt", "path")).toBe("a/b/file.txt");
    });

    it("allows a hidden-style filename like '.gitignore'", () => {
      expect(validateDrivePath(".gitignore", "path")).toBe(".gitignore");
    });

    it("allows an empty path", () => {
      expect(validateDrivePath("", "path")).toBe("");
    });
  });

  describe("illegal characters", () => {
    it("throws when containing a question mark", () => {
      const err = capture(() => validateDrivePath("a?b", "path"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("path");
    });

    it("throws when containing a hash", () => {
      const err = capture(() => validateDrivePath("a#b", "path"));
      expect(err.status).toBe(400);
    });

    it("throws when containing a colon", () => {
      const err = capture(() => validateDrivePath("a:b", "path"));
      expect(err.status).toBe(400);
    });

    it("throws when containing a NUL byte", () => {
      const err = capture(() => validateDrivePath("a\x00b", "path"));
      expect(err.status).toBe(400);
    });
  });

  describe("path traversal", () => {
    it("throws on a '..' segment in the middle", () => {
      const err = capture(() => validateDrivePath("a/../b", "path"));
      expect(err.status).toBe(400);
      expect(err.message).toContain("path");
    });

    it("throws on a '.' segment in the middle", () => {
      const err = capture(() => validateDrivePath("a/./b", "path"));
      expect(err.status).toBe(400);
    });

    it("throws on a lone '..'", () => {
      const err = capture(() => validateDrivePath("..", "path"));
      expect(err.status).toBe(400);
    });

    it("throws on a lone '.'", () => {
      const err = capture(() => validateDrivePath(".", "path"));
      expect(err.status).toBe(400);
    });

    it("throws on a leading '..' segment", () => {
      const err = capture(() => validateDrivePath("../b", "path"));
      expect(err.status).toBe(400);
    });

    it("throws on a trailing '..' segment", () => {
      const err = capture(() => validateDrivePath("a/..", "path"));
      expect(err.status).toBe(400);
    });
  });
});

describe("escapeKql", () => {
  it("escapes backslash THEN double-quote (order matters)", () => {
    // Input: a\"b  ->  backslash doubled first (a\\"b), then quote escaped.
    expect(escapeKql('a\\"b')).toBe('a\\\\\\"b');
  });

  it("doubles a single backslash", () => {
    expect(escapeKql("a\\b")).toBe("a\\\\b");
  });

  it("escapes a lone double quote", () => {
    expect(escapeKql('a"b')).toBe('a\\"b');
  });

  it("leaves a plain string unchanged", () => {
    expect(escapeKql("hello world")).toBe("hello world");
  });

  it("leaves an empty string unchanged", () => {
    expect(escapeKql("")).toBe("");
  });

  it("escapes multiple occurrences", () => {
    expect(escapeKql('"\\"')).toBe('\\"\\\\\\"');
  });
});

describe("escapeODataString", () => {
  it("doubles a single quote", () => {
    expect(escapeODataString("O'Brien")).toBe("O''Brien");
  });

  it("leaves an empty string unchanged", () => {
    expect(escapeODataString("")).toBe("");
  });

  it("leaves a string without single quotes unchanged", () => {
    expect(escapeODataString("hello")).toBe("hello");
  });

  it("doubles multiple single quotes", () => {
    expect(escapeODataString("'a'b'")).toBe("''a''b''");
  });

  it("handles consecutive single quotes", () => {
    expect(escapeODataString("a''b")).toBe("a''''b");
  });
});
