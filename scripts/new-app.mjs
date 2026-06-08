#!/usr/bin/env node
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const name = process.argv[2];

if (!name) {
  console.error("usage: pnpm new-app <name>");
  console.error("  <name> must be lowercase, e.g. 'todo' or 'one-note'");
  process.exit(1);
}
if (!/^[a-z][a-z0-9-]*$/.test(name)) {
  console.error(`invalid name '${name}'. Use lowercase letters/digits/hyphens: [a-z][a-z0-9-]*`);
  process.exit(1);
}

const appDir = path.join(ROOT, "apps", name);
if (existsSync(appDir)) {
  console.error(`apps/${name} already exists. Aborting.`);
  process.exit(1);
}

const title = name
  .split("-")
  .map((s) => s.charAt(0).toUpperCase() + s.slice(1))
  .join(" ");
const serverName = `microsoft-${name}`;

mkdirSync(path.join(appDir, "src"), { recursive: true });

const packageJson = {
  name: `@microsoft-mcp/${name}`,
  version: "1.0.0",
  type: "module",
  description: `Microsoft ${title} MCP server.`,
  license: "MIT",
  author: "0xka13b",
  homepage: `https://github.com/0xka13b/microsoft-mcps/tree/main/apps/${name}#readme`,
  repository: {
    type: "git",
    url: "git+https://github.com/0xka13b/microsoft-mcps.git",
    directory: `apps/${name}`,
  },
  bugs: { url: "https://github.com/0xka13b/microsoft-mcps/issues" },
  keywords: ["mcp", "model-context-protocol", "microsoft", "microsoft-365", "microsoft-graph", name],
  bin: { [`microsoft-${name}-mcp`]: "./dist/index.js" },
  files: ["dist"],
  engines: { node: ">=20" },
  publishConfig: { access: "public" },
  scripts: {
    dev: "tsx watch src/index.ts",
    start: "node dist/index.js",
    build: "tsup",
    "check-types": "tsc --noEmit",
    clean: "rm -rf dist .turbo",
    prepublishOnly: "tsup",
  },
  // @microsoft-mcp/* are bundled into dist by tsup, so they are build-time only.
  dependencies: {
    "@modelcontextprotocol/sdk": "^1.29.0",
    express: "^5.2.1",
    zod: "^4.4.3",
  },
  devDependencies: {
    "@microsoft-mcp/core": "workspace:*",
    "@microsoft-mcp/graph": "workspace:*",
    "@microsoft-mcp/validation": "workspace:*",
    "@types/node": "^22.10.0",
    tsup: "^8.3.5",
    tsx: "^4.19.2",
    typescript: "^5.8.2",
  },
};

const tsconfig = { extends: "../../tsconfig.base.json", include: ["src/**/*"] };

const tsupConfig = `import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  target: "node20",
  platform: "node",
  outDir: "dist",
  clean: true,
  bundle: true,
  noExternal: [/^@microsoft-mcp\\//],
  banner: { js: "#!/usr/bin/env node" },
  sourcemap: true,
});
`;

const toolsTs = `import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";

export const tools = [
  defineTool({
    name: "me",
    description: "Get the signed-in user's profile (id, displayName, mail).",
    inputSchema: {},
    confirmationPolicy: "never",
    handler: ({ graph }) =>
      graph.request("GET", "/me", undefined, { $select: "id,displayName,mail" }),
  }),
];
`;

const indexTs = `import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "${serverName}", version: "1.0.0", title: "Microsoft ${title}" }, tools);
`;

writeFileSync(path.join(appDir, "package.json"), JSON.stringify(packageJson, null, 2) + "\n");
writeFileSync(path.join(appDir, "tsconfig.json"), JSON.stringify(tsconfig, null, 2) + "\n");
writeFileSync(path.join(appDir, "tsup.config.ts"), tsupConfig);
writeFileSync(path.join(appDir, "src", "tools.ts"), toolsTs);
writeFileSync(path.join(appDir, "src", "index.ts"), indexTs);

console.log(`✓ scaffolded apps/${name}/`);
console.log("");
console.log("next steps:");
console.log("  pnpm install");
console.log(`  pnpm --filter @microsoft-mcp/${name} dev`);
console.log("");
console.log(`tools live in apps/${name}/src/tools.ts`);
