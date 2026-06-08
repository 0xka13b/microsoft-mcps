import { defineConfig } from "tsup";

// Bundles the app and its workspace packages (@microsoft-mcp/*) into a single
// ESM entrypoint with a node shebang, runnable as `node dist/index.js` or via
// the `bin`. Third-party deps (SDK, express, zod) stay external and resolve
// from node_modules at runtime.
export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  target: "node20",
  platform: "node",
  outDir: "dist",
  clean: true,
  bundle: true,
  noExternal: [/^@microsoft-mcp\//],
  banner: { js: "#!/usr/bin/env node" },
  sourcemap: true,
});
