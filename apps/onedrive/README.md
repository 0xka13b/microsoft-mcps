# ms-onedrive-mcp

Microsoft OneDrive [Model Context Protocol](https://modelcontextprotocol.io) server — 9 tools over the [Microsoft Graph](https://learn.microsoft.com/graph/) `v1.0` API. Runs over **stdio** (local) or **Streamable HTTP** (remote).

Part of [microsoft-mcps](https://github.com/0xka13b/microsoft-mcps).

## Setup

Sign in once with your own free Microsoft [Entra ID](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app) app. Register an app as a **public client** (redirect URI `http://localhost`, "Allow public client flows" = Yes) and grant these **delegated** Microsoft Graph permissions: `User.Read`, `Files.ReadWrite.All`. Then:

```bash
export MICROSOFT_CLIENT_ID=<your-client-id>
npx -y ms-onedrive-mcp login              # opens the browser
# npx -y ms-onedrive-mcp login --device-code   # headless alternative
```

The refresh token is cached under `~/.config/microsoft-mcp/` and renewed automatically — no token pasting, no 1-hour expiry.

## Use over stdio (e.g. Claude Desktop)

```jsonc
{
  "mcpServers": {
    "microsoft-onedrive": {
      "command": "npx",
      "args": ["-y", "ms-onedrive-mcp"],
      "env": { "MICROSOFT_CLIENT_ID": "<your-client-id>" }
    }
  }
}
```

Advanced: set `MICROSOFT_ACCESS_TOKEN` to supply a pre-acquired Graph token instead of signing in.

## Remote (Streamable HTTP)

```bash
npx -y ms-onedrive-mcp --http --port 3000
```

Point a Streamable-HTTP MCP client at `http://localhost:3000/mcp` and send the token per request as `Authorization: Bearer <token>`.

See the [main repo](https://github.com/0xka13b/microsoft-mcps#readme) for the full Entra app walkthrough, environment variables, and the tool list.

## License

MIT
