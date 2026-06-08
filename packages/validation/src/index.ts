/**
 * Input validation and sanitization shared across the Microsoft MCP servers.
 *
 * These guard against path traversal (IDs interpolated into Graph URL segments)
 * and query injection (user input interpolated into KQL / OData string literals).
 */

/**
 * Creates an Error carrying an HTTP `status`. The core tool wrapper reads
 * `status` to map failures onto MCP error results with the right semantics
 * (400 client error, 401/403 auth, etc.).
 */
export const httpError = (message: string, status: number): Error & { status: number } =>
  Object.assign(new Error(message), { status });

/**
 * Rejects resource IDs that contain path-unsafe characters.
 * Prevents path traversal when IDs are interpolated into URL segments.
 * Also validates Microsoft Graph ID format.
 */
export const validateId = (id: string, label: string): string => {
  if (!id || typeof id !== "string") {
    throw httpError(`Invalid ${label}: ID is required`, 400);
  }

  // Path-unsafe characters.
  if (/[/?#\x00]/.test(id)) {
    throw httpError(`Invalid ${label}: contains illegal characters`, 400);
  }

  // Graph IDs never contain whitespace.
  if (id.includes(" ") || id.includes("\t") || id.includes("\n") || id.includes("\r")) {
    throw httpError(`Invalid ${label}: ID contains whitespace`, 400);
  }

  // Common malformed values from upstream serialization bugs.
  if (id.toLowerCase() === "undefined" || id.toLowerCase() === "null" || id === "[object Object]") {
    throw httpError(`Invalid ${label}: malformed ID`, 400);
  }

  return id;
};

/**
 * Rejects drive paths that contain `.` or `..` segments.
 * Prevents path traversal in OneDrive / SharePoint path-based API calls.
 */
export const validateDrivePath = (path: string, label: string): string => {
  if (/[?#:\x00]/.test(path)) {
    throw httpError(`Invalid ${label}: illegal character`, 400);
  }
  if (path.split("/").some((seg) => seg === ".." || seg === ".")) {
    throw httpError(`Invalid ${label}: path traversal not allowed`, 400);
  }
  return path;
};

/**
 * Escapes double quotes in KQL / OData `$search` string literals.
 * Prevents query injection when user input is interpolated as `$search: "<query>"`.
 */
export const escapeKql = (q: string): string =>
  q.replace(/\\/g, "\\\\").replace(/"/g, '\\"');

/**
 * Escapes single quotes in OData function-call string literals (e.g. `search(q='...')`).
 * OData convention: a literal `'` is represented as `''`.
 */
export const escapeODataString = (q: string): string => q.replace(/'/g, "''");
