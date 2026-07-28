import { build } from "esbuild";
import { mkdir, rm } from "node:fs/promises";
import { resolve } from "node:path";
import { spawnSync } from "node:child_process";

const root = resolve(import.meta.dirname, "..");
const out = resolve(root, ".test-build");
await rm(out, { recursive: true, force: true });
await mkdir(out, { recursive: true });

await build({
  entryPoints: [resolve(root, "tests/core.test.ts")],
  bundle: true,
  platform: "node",
  format: "esm",
  target: "node20",
  outfile: resolve(out, "core.test.mjs"),
});

const result = spawnSync(process.execPath, ["--test", resolve(out, "core.test.mjs")], {
  cwd: root,
  stdio: "inherit",
});
process.exit(result.status ?? 1);
