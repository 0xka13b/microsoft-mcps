// Reads the vitest coverage summary and writes a shields.io "endpoint" badge
// JSON (consumed by the coverage badge in the README). Run after `pnpm
// test:coverage`. CI regenerates and commits it on pushes to master.
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";

const summary = JSON.parse(readFileSync("coverage/coverage-summary.json", "utf8"));
const pct = Math.round(summary.total.lines.pct);

const color =
  pct >= 90 ? "brightgreen" :
  pct >= 80 ? "green" :
  pct >= 70 ? "yellowgreen" :
  pct >= 60 ? "yellow" :
  pct >= 50 ? "orange" :
  "red";

mkdirSync(".github/badges", { recursive: true });
writeFileSync(
  ".github/badges/coverage.json",
  JSON.stringify({ schemaVersion: 1, label: "coverage", message: `${pct}%`, color }) + "\n",
);

console.log(`coverage badge -> ${pct}% (${color})`);
