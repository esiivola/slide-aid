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

const wrappers = `
/** @OnlyCurrentDoc */
function onOpen(e) { return SlideAidBundle.onOpen(e); }
function onInstall(e) { return SlideAidBundle.onInstall(e); }
function showSlideAidSidebar() { return SlideAidBundle.showSlideAidSidebar(); }
function menuBuildColumn() { return SlideAidBundle.menuBuildColumn(); }
function getSidebarState() { return SlideAidBundle.getSidebarState(); }
function setReferenceFromSelection() { return SlideAidBundle.setReferenceFromSelection(); }
function clearReference() { return SlideAidBundle.clearReference(); }
function runSlideAidCommand(request) { return SlideAidBundle.runSlideAidCommand(request); }
function buildSlideAidChart(kind) { return SlideAidBundle.buildSlideAidChart(kind); }
function rebuildSlideAidChart() { return SlideAidBundle.rebuildSlideAidChart(); }
function editSlideAidChartData() { return SlideAidBundle.editSlideAidChartData(); }
function restyleSlideAidCharts(allCharts) { return SlideAidBundle.restyleSlideAidCharts(allCharts); }
function setSlideAidPalette(name) { return SlideAidBundle.setSlideAidPalette(name); }
function recolorSlideAidSeries(seriesIndex, color) { return SlideAidBundle.recolorSlideAidSeries(seriesIndex, color); }
function buildLinkedSlideAidChart(kind, spreadsheetUrl, sheetName, rangeA1) { return SlideAidBundle.buildLinkedSlideAidChart(kind, spreadsheetUrl, sheetName, rangeA1); }
function validateLinkedSlideAidChart(kind, spreadsheetUrl, sheetName, rangeA1) { return SlideAidBundle.validateLinkedSlideAidChart(kind, spreadsheetUrl, sheetName, rangeA1); }
function refreshLinkedSlideAidChart() { return SlideAidBundle.refreshLinkedSlideAidChart(); }
function applySlideAidThemeColor(target, themeName) { return SlideAidBundle.applySlideAidThemeColor(target, themeName); }
function convertSlideAidColors(toTheme) { return SlideAidBundle.convertSlideAidColors(toTheme); }
function applySlideAidPaletteToTheme(paletteName) { return SlideAidBundle.applySlideAidPaletteToTheme(paletteName); }
function saveSlideAidLayout(name) { return SlideAidBundle.saveSlideAidLayout(name); }
function applySlideAidLayout(name) { return SlideAidBundle.applySlideAidLayout(name); }
function deleteSlideAidLayout(name) { return SlideAidBundle.deleteSlideAidLayout(name); }
function configureSlideAidLibrary(url) { return SlideAidBundle.configureSlideAidLibrary(url); }
function getSlideAidLibraryItems() { return SlideAidBundle.getSlideAidLibraryItems(); }
function insertSlideAidLibraryItem(slideId) { return SlideAidBundle.insertSlideAidLibraryItem(slideId); }
function addSelectionToSlideAidLibrary(name) { return SlideAidBundle.addSelectionToSlideAidLibrary(name); }
function refreshSelectedSlideAidLibraryItem() { return SlideAidBundle.refreshSelectedSlideAidLibraryItem(); }
function insertSlideAidIcon(icon, color) { return SlideAidBundle.insertSlideAidIcon(icon, color); }
function scanSlideAidDeck() { return SlideAidBundle.scanSlideAidDeck(); }
function focusSlideAidQaIssue(slideId, objectId) { return SlideAidBundle.focusSlideAidQaIssue(slideId, objectId); }
function fixSlideAidQaIssue(issue) { return SlideAidBundle.fixSlideAidQaIssue(issue); }
`;

const codePath = resolve(dist, "Code.js");
const code = await readFile(codePath, "utf8");
await writeFile(codePath, `${code}\n${wrappers}`, "utf8");
let sidebar = await readFile(resolve(root, "src/ui/Sidebar.html"), "utf8");
const iconAidCatalog = JSON.parse(await readFile(resolve(root, "../../shared/iconaid/catalog.json"), "utf8"));
sidebar = sidebar.replace(
  "{{ICONAID_CATALOG}}",
  JSON.stringify(iconAidCatalog).replaceAll("<", "\\u003c"),
);
const iconNames = [...sidebar.matchAll(/\{\{ICON:([a-z0-9_]+)\}\}/gi)].map((match) => match[1]);
for (const iconName of new Set(iconNames)) {
  const bytes = await readFile(resolve(root, "../../shared/icons", `${iconName}.png`));
  sidebar = sidebar.replaceAll(`{{ICON:${iconName}}}`, `data:image/png;base64,${bytes.toString("base64")}`);
}
await writeFile(resolve(dist, "Sidebar.html"), sidebar, "utf8");
await cp(resolve(root, "appsscript.json"), resolve(dist, "appsscript.json"));

console.log("Built dist/Code.js, dist/Sidebar.html and dist/appsscript.json");
