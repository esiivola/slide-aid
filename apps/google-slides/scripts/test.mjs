import { build } from "esbuild";
import { mkdir, rm } from "node:fs/promises";
import { resolve } from "node:path";
import { spawnSync } from "node:child_process";

const root = resolve(import.meta.dirname, "..");
const out = resolve(root, ".test-build");
await rm(out, { recursive: true, force: true });
await mkdir(out, { recursive: true });

// core.test covers the pure logic; host.test drives the Apps Script entry points
// against a fake host so the request-building paths are exercised too.
await build({
  entryPoints: [resolve(root, "tests/core.test.ts"), resolve(root, "tests/host.test.ts")],
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node20",
  outdir: out,
  outExtension: { ".js": ".mjs" },
});

const result = spawnSync(
  process.execPath,
  ["--test", resolve(out, "core.test.mjs"), resolve(out, "host.test.mjs")],
  { cwd: root, stdio: "inherit" },
);
process.exit(result.status ?? 1);
