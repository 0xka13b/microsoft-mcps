/**
 * Structured JSON logger for the Microsoft MCP servers.
 *
 * IMPORTANT: every level writes to **stderr**, never stdout. When a server runs
 * over the stdio transport, stdout is reserved exclusively for the JSON-RPC
 * message stream — writing logs there would corrupt the protocol. stderr is
 * always safe and is what MCP clients surface as server logs.
 */
export type LogLevel = "debug" | "info" | "warn" | "error";

type LogData = Record<string, unknown>;

const emit = (level: LogLevel, msg: string, data: LogData): void => {
  const line = JSON.stringify({
    ts: new Date().toISOString(),
    level,
    msg,
    ...data,
  });
  process.stderr.write(line + "\n");
};

export const log = {
  /** Verbose diagnostics; only emitted when MCP_DEBUG is set. */
  debug: (msg: string, data: LogData = {}): void => {
    if (process.env.MCP_DEBUG) emit("debug", msg, data);
  },
  info: (msg: string, data: LogData = {}): void => emit("info", msg, data),
  warn: (msg: string, data: LogData = {}): void => emit("warn", msg, data),
  error: (msg: string, data: LogData = {}): void => emit("error", msg, data),
};

export type Logger = typeof log;
