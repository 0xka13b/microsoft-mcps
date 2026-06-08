# Microsoft MCP

Microsoft 365 [Model Context Protocol](https://modelcontextprotocol.io) servers — **Calendar, Contacts, OneDrive, Outlook, and SharePoint** — built on the official [`@modelcontextprotocol/sdk`](https://www.npmjs.com/package/@modelcontextprotocol/sdk).

Each server speaks the real MCP protocol and runs over **either transport**:

- **stdio** — for local MCP clients that launch the server as a subprocess (Claude Desktop, IDEs, the MCP Inspector).
- **Streamable HTTP** — for remote/hosted use, with the Microsoft Graph access token supplied per request via `Authorization: Bearer`.

This is a from-scratch reimplementation of an earlier set of Hono-based REST services. Every tool's behavior (schemas, Graph endpoints, edge cases) is preserved; only the framing is now standard MCP.

## Servers

| Server      | Package                      | Tools | Binary                       |
| ----------- | ---------------------------- | ----: | ---------------------------- |
| Calendar    | `@microsoft-mcp/calendar`    |     9 | `microsoft-calendar-mcp`     |
| Contacts    | `@microsoft-mcp/contacts`    |     7 | `microsoft-contacts-mcp`     |
| OneDrive    | `@microsoft-mcp/onedrive`    |     9 | `microsoft-onedrive-mcp`     |
| Outlook     | `@microsoft-mcp/outlook`     |    14 | `microsoft-outlook-mcp`      |
| SharePoint  | `@microsoft-mcp/sharepoint`  |    23 | `microsoft-sharepoint-mcp`   |

All tools are thin wrappers over the [Microsoft Graph](https://learn.microsoft.com/graph/) `v1.0` API.

## Layout

```
microsoft-mcp/
├── apps/                       # one MCP server per Microsoft 365 product
│   ├── calendar/
│   ├── contacts/
│   ├── onedrive/
│   ├── outlook/
│   └── sharepoint/
│       └── src/
│           ├── tools.ts        # declarative tool definitions (schema + handler)
│           └── index.ts        # run({ name, version }, tools)
└── packages/                   # shared building blocks
    ├── core/                   # MCP server bootstrap + dual transport (stdio / HTTP)
    ├── graph/                  # Microsoft Graph HTTP client
    ├── validation/             # id / path / query sanitizers
    └── logger/                 # structured JSON logging (stderr-only — stdio-safe)
```

A server is just a list of tools handed to `run()`:

```ts
// apps/calendar/src/index.ts
import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "microsoft-calendar", version: "1.0.0", title: "Microsoft Calendar" }, tools);
```

```ts
// a single tool
defineTool({
  name: "get_event",
  description: "Get a single calendar event by ID.",
  inputSchema: { event_id: z.string().describe("Event ID") },
  confirmationPolicy: "never",
  handler: ({ graph }, { event_id }) => {
    validateId(event_id, "event_id");
    return graph.request("GET", `/me/events/${event_id}`);
  },
});
```

`confirmationPolicy` (`"always"` for mutating/destructive tools, `"never"` for read-only) is surfaced to clients as MCP `readOnlyHint` / `destructiveHint` annotations.

## Requirements

- Node.js >= 20
- pnpm 10 (`corepack enable`)

## Setup

```bash
pnpm install
pnpm build        # build all servers (turbo) -> apps/*/dist/index.js
pnpm check-types  # typecheck everything
```

## Authentication

These servers do **not** run an OAuth flow; supply a pre-acquired Microsoft Graph access token.

- **stdio:** set `MICROSOFT_ACCESS_TOKEN` in the environment.
- **HTTP:** send `Authorization: Bearer <token>` on each `POST /mcp` request. Each request is handled statelessly with its own token, so multiple callers never share credentials.

For local testing you can mint a short-lived token with the Azure CLI:

```bash
az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv
```

Copy `.env.example` to `.env` and fill in the values for local development.

## Running

### stdio (e.g. Claude Desktop)

Each server is published to npm and runnable with `npx` — no clone or build:

```jsonc
// claude_desktop_config.json
{
  "mcpServers": {
    "microsoft-calendar": {
      "command": "npx",
      "args": ["-y", "@microsoft-mcp/calendar"],
      "env": { "MICROSOFT_ACCESS_TOKEN": "<token>" }
    }
  }
}
```

Or point at a local build instead of npm:

```jsonc
{
  "command": "node",
  "args": ["/abs/path/microsoft-mcp/apps/calendar/dist/index.js"],
  "env": { "MICROSOFT_ACCESS_TOKEN": "<token>" }
}
```

During development you can skip the build and run the TypeScript directly:

```bash
MICROSOFT_ACCESS_TOKEN=<token> pnpm --filter @microsoft-mcp/calendar dev
```

### Streamable HTTP

```bash
# build first, then:
PORT=3000 node apps/calendar/dist/index.js --http
# or, in dev:
pnpm --filter @microsoft-mcp/calendar dev -- --http --port 3000
```

The server exposes `POST /mcp` (the MCP endpoint) and `GET /healthz`. Point any Streamable-HTTP MCP client at `http://localhost:3000/mcp` with an `Authorization: Bearer` header.

## Transport selection

Resolved in this order: `--stdio` / `--http` flag → `MCP_TRANSPORT=stdio|http` → default `stdio`.
HTTP port: `--port <n>` → `PORT` → `3000`.

## Environment variables

| Variable                  | Used by | Description                                                        |
| ------------------------- | ------- | ------------------------------------------------------------------ |
| `MICROSOFT_ACCESS_TOKEN`  | stdio   | Microsoft Graph access token.                                      |
| `MCP_TRANSPORT`           | both    | `stdio` (default) or `http`.                                       |
| `PORT`                    | http    | Listen port (default `3000`).                                      |
| `MCP_HTTP_BODY_LIMIT`     | http    | Max request body size (default `50mb`) for base64 uploads.         |
| `MCP_DEBUG`               | both    | Any non-empty value enables debug logs (to stderr).                |

## Adding a server

```bash
pnpm new-app <name>
```

Scaffolds `apps/<name>/` from the standard template (package.json, tsconfig, tsup config, and a starter `src/tools.ts` + `src/index.ts`). Then add your tools and run `pnpm install`. The template is already publish-ready.

## Releasing

Each server is an independent npm package. The shared `packages/*` are `private` — `tsup` bundles them into each server's `dist/index.js` (`noExternal: [/^@microsoft-mcp\//]`), so only `@modelcontextprotocol/sdk`, `express`, and `zod` are installed at runtime.

Versioning and publishing use [changesets](https://github.com/changesets/changesets):

```bash
pnpm changeset           # describe a change; pick affected servers + bump type
pnpm version-packages    # apply pending changesets -> bump versions + changelogs
pnpm release             # build all, then `changeset publish` to npm
```

`pnpm release` runs `prepublishOnly` (a fresh `tsup` build) per package, so the published tarball always contains an up-to-date bundle. Packages publish with `--access public`. Publishing requires `npm login` (an npm account with access to the `@microsoft-mcp` scope).
