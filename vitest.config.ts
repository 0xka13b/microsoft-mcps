import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["packages/**/*.test.ts", "apps/**/*.test.ts"],
    environment: "node",
    coverage: {
      provider: "v8",
      reportsDirectory: "./coverage",
      reporter: ["text", "text-summary", "json-summary", "lcov", "html"],
      // Coverage targets the shared library code, where the logic lives. The
      // apps/ packages are declarative Graph tool definitions, exercised by the
      // per-app invariant suites (apps/*/src/tools.test.ts) rather than measured
      // for line coverage.
      include: ["packages/*/src/**/*.ts"],
      exclude: ["**/*.test.ts", "**/dist/**"],
      // Fail CI if coverage of the shared library code regresses below this.
      thresholds: {
        lines: 85,
        statements: 85,
        functions: 85,
        branches: 85,
      },
    },
  },
});
