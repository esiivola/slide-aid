import { build } from "esbuild";
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const dist = resolve(root, "dist");

await rm(dist, { recursive: true, force: true });
await mkdir(dist, { recursive: true });

await build({
  entryPoints: [resolve(root, "src/entrypoints/index.ts")],
  bundle: true,
  format: "iife",
  globalName: "SlideAidBundle",
  target: "es2020",
  outfile: resolve(dist, "Code.js"),
  legalComments: "none",
  minify: false,
});

// Apps Script can only call top-level globals, so every exported entrypoint needs
// a thin forwarding wrapper. These are derived from the source rather than
// hand-listed: a hand-list silently drops new entrypoints, and the sidebar's only
// symptom is a "function not found" at click time.
const entrypointSource = await readFile(resolve(root, "src/entrypoints/index.ts"), "utf8");
const exported = [...entrypointSource.matchAll(/^export function ([A-Za-z0-9_]+)\s*\(/gm)].map((match) => match[1]);
if (!exported.length) throw new Error("No entrypoints found in src/entrypoints/index.ts.");
const duplicates = exported.filter((name, index) => exported.indexOf(name) !== index);
if (duplicates.length) throw new Error(`Duplicate entrypoints: ${duplicates.join(", ")}`);

const wrappers = [
  "/** @OnlyCurrentDoc */",
  ...exported.map((name) => `function ${name}() { return SlideAidBundle.${name}.apply(null, arguments); }`),
].join("\n");

const codePath = resolve(dist, "Code.js");
const code = await readFile(codePath, "utf8");
await writeFile(codePath, `${code}\n${wrappers}`, "utf8");
let sidebar = await readFile(resolve(root, "src/ui/Sidebar.html"), "utf8");
const iconAidCatalog = JSON.parse(await readFile(resolve(root, "../../shared/iconaid/catalog.json"), "utf8"));
sidebar = sidebar.replace(
  "{{ICONAID_CATALOG}}",
  JSON.stringify(iconAidCatalog).replaceAll("<", "\\u003c"),
);

// The full ~10k catalog is split: metadata rides in the sidebar so search is
// instant, while the path data ships as numbered project files that the server
// reads on demand. Apps Script stores only .gs and .html, so each JSON payload
// travels inside an .html file.
const slidesIconDir = resolve(root, "../../shared/iconaid/slides");
const iconIndex = JSON.parse(await readFile(resolve(slidesIconDir, "index.json"), "utf8"));
sidebar = sidebar.replace(
  "{{ICON_INDEX}}",
  JSON.stringify(iconIndex).replaceAll("<", "\\u003c"),
);
await writeFile(resolve(dist, "IconShards.html"), JSON.stringify(iconIndex.shards), "utf8");
for (const shard of iconIndex.shards) {
  const name = `paths-${String(shard.shard).padStart(2, "0")}.json`;
  const payload = await readFile(resolve(slidesIconDir, name), "utf8");
  if (payload.includes("<")) throw new Error(`${name} contains a character that is unsafe inside an HTML project file.`);
  await writeFile(resolve(dist, `IconPaths${String(shard.shard).padStart(2, "0")}.html`), payload, "utf8");
}
console.log(`Embedded ${iconIndex.icons.length} icon index entries and ${iconIndex.shards.length} path shards.`);
const iconNames = [...sidebar.matchAll(/\{\{ICON:([a-z0-9_]+)\}\}/gi)].map((match) => match[1]);
for (const iconName of new Set(iconNames)) {
  const bytes = await readFile(resolve(root, "../../shared/icons", `${iconName}.png`));
  sidebar = sidebar.replaceAll(`{{ICON:${iconName}}}`, `data:image/png;base64,${bytes.toString("base64")}`);
}
await writeFile(resolve(dist, "Sidebar.html"), sidebar, "utf8");
await cp(resolve(root, "appsscript.json"), resolve(dist, "appsscript.json"));

console.log("Built dist/Code.js, dist/Sidebar.html and dist/appsscript.json");
