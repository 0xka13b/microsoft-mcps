# @microsoft-mcp/contacts

Microsoft Contacts [Model Context Protocol](https://modelcontextprotocol.io) server — 7 tools over the [Microsoft Graph](https://learn.microsoft.com/graph/) `v1.0` API. Runs over **stdio** (local) or **Streamable HTTP** (remote).

Part of [microsoft-mcps](https://github.com/0xka13b/microsoft-mcps).

## Install (stdio)

Add to your MCP client config (e.g. Claude Desktop `claude_desktop_config.json`):

```jsonc
{
  "mcpServers": {
    "microsoft-contacts": {
      "command": "npx",
      "args": ["-y", "@microsoft-mcp/contacts"],
      "env": { "MICROSOFT_ACCESS_TOKEN": "<microsoft-graph-token>" }
    }
  }
}
```

These servers do **not** run an OAuth flow — supply a pre-acquired Microsoft Graph access token. For local testing you can mint one with the Azure CLI:

```bash
az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv
```

## Remote (Streamable HTTP)

```bash
npx -y @microsoft-mcp/contacts --http --port 3000
```

Point a Streamable-HTTP MCP client at `http://localhost:3000/mcp` and send the token per request as `Authorization: Bearer <token>`.

See the [main repo](https://github.com/0xka13b/microsoft-mcps#readme) for transport selection, environment variables, and the full tool list.

## License

MIT
