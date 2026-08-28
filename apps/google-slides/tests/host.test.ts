import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { installHost, insertedText, requestsOfKind, restoreHost, type HostState } from "./host";
import { insertEditableIcon, insertIconImage, makeIconsEditable } from "../src/slides/icons";
import { flattenPath, polylineSegments, type IconPolyline } from "../src/core/icon-path";
import { fillSpans, isFilledIcon } from "../src/core/icon-fill";
import { buildChart, applyChartSettings, chartSettingsState, saveFamilyPalette } from "../src/charts/charts";
import { annotateDifference, insertHarveyBall } from "../src/charts/annotations";
import { barTag } from "../src/core/chart-data";
import { iconPathsFor } from "../src/slides/icon-catalog";
import { hideObjects, unhideAll } from "../src/commands/deck-commands";
import { snapToTable } from "../src/commands/object-commands";
import { executeCommand } from "../src/commands/geometry-commands";
import { addReviewCallout, addReviewNote, initialsFrom, removeReviewMarkup, setReviewInitials } from "../src/slides/review";

// A tiny two-shard catalog: one stroke icon and one solid one.
const STROKE_PATH = "M4 4 L20 4 L20 20 L4 20 Z";
const SOLID_PATH = "M2 2 L22 2 L22 22 L2 22 Z";
const PROJECT_FILES = {
  IconShards: JSON.stringify([
    { shard: 0, firstId: "bootstrap-box", lastId: "bootstrap-box" },
    { shard: 1, firstId: "lucide-frame", lastId: "lucide-frame" },
  ]),
  // Duplicate paths mirror a defect present in hundreds of upstream catalog
  // entries. Production must normalize these before even-odd fill conversion.
  IconPaths00: JSON.stringify({ "bootstrap-box": [SOLID_PATH, SOLID_PATH] }),
  IconPaths01: JSON.stringify({ "lucide-frame": [STROKE_PATH] }),
};

function teardown(): void {
  restoreHost();
}

test("inserting an icon stores a tagged picture, not shapes", (t) => {
  t.after(teardown);
  const state = installHost();
  const result = insertIconImage("lucide-frame", "Frame", "#1F497D", "QUJD");

  assert.equal(result.ok, true);
  const images = [...state.elements.values()].filter((item) => item.type === "IMAGE");
  assert.equal(images.length, 1, "one picture, matching the PowerPoint task pane");
  assert.equal(state.batchUpdates, 0, "no shape batch is sent for a picture insert");
  assert.equal(images[0]!.image?.contentType, "image/png");
  assert.equal(images[0]!.title, "IconAid: Frame");
  // The tag is what Make Editable and PowerPoint both key off.
  assert.match(images[0]!.description, /\[iconaid:lucide-frame:#1F497D\]/);
  // Centred on a 720x405 page at 72pt.
  assert.equal(images[0]!.left, 324);
  assert.equal(images[0]!.top, 166.5);
});

test("icon insertion rejects malformed input", (t) => {
  t.after(teardown);
  installHost();
  assert.throws(() => insertIconImage("Bad Id", "x", "#1F497D", "QUJD"), /icon id is invalid/);
  assert.throws(() => insertIconImage("lucide-frame", "x", "red", "QUJD"), /#RRGGBB/);
  assert.throws(() => insertIconImage("lucide-frame", "x", "#1F497D", "not base64!"), /could not be read/);
});

test("direct icon insertion creates an editable stroke group", (t) => {
  t.after(teardown);
  const state = installHost({ projectFiles: PROJECT_FILES });

  const result = insertEditableIcon("lucide-frame", "Frame", "#123456");

  assert.match(result.message, /Inserted Frame as 4 editable shapes/);
  assert.equal(state.batchUpdates, 1, "the editable icon is inserted atomically");
  assert.equal(requestsOfKind(state, "createImage").length, 0, "no picture is inserted");
  assert.equal(requestsOfKind(state, "createLine").length, 4, "the closed path becomes native lines");
  assert.equal(requestsOfKind(state, "createShape").length, 0, "a stroke icon has no filled slices");
  const group = requestsOfKind(state, "groupObjects")[0]!;
  assert.equal((group.childrenObjectIds as string[]).length, 4);
  const alt = requestsOfKind(state, "updatePageElementAltText").at(-1)!;
  assert.match(String(alt.description), /\[iconaid:lucide-frame:#123456\]/);
});

test("direct icon insertion creates an editable solid group", (t) => {
  t.after(teardown);
  const state = installHost({ projectFiles: PROJECT_FILES });

  const result = insertEditableIcon("bootstrap-box", "Box", "#000000");

  assert.match(result.message, /Inserted Box as .* editable shapes/);
  assert.equal(state.batchUpdates, 1);
  assert.equal(requestsOfKind(state, "createLine").length, 0, "solid icons stay filled");
  assert.ok(requestsOfKind(state, "createShape").length > 20, "the filled path becomes native slices");
});

test("direct icon insertion validates ids and deployment data", (t) => {
  t.after(teardown);
  installHost({ projectFiles: PROJECT_FILES });
  assert.throws(() => insertEditableIcon("Bad Id", "x", "#1F497D"), /icon id is invalid/);
  assert.throws(() => insertEditableIcon("tabler-unknown", "x", "#1F497D"), /not in this deployment's catalog/);
});

test("all 54,250 catalog icons produce editable Slides geometry", () => {
  const slidesDir = resolve(process.cwd(), "../../shared/iconaid/slides");
  const index = JSON.parse(readFileSync(resolve(slidesDir, "index.json"), "utf8")) as {
    icons: { id: string; k: number; f?: number }[];
  };
  const shards = new Map<number, Record<string, string[]>>();
  for (const icon of index.icons) {
    let shard = shards.get(icon.k);
    if (!shard) {
      shard = JSON.parse(readFileSync(resolve(slidesDir, `paths-${String(icon.k).padStart(2, "0")}.json`), "utf8")) as Record<string, string[]>;
      shards.set(icon.k, shard);
    }
    const paths = shard[icon.id];
    assert.ok(paths?.length, `${icon.id} has no path data`);
    const runs: IconPolyline[] = [];
    for (const path of new Set(paths)) {
      try {
        runs.push(...flattenPath(path));
      } catch {
        // Production skips isolated malformed subpaths and preserves the rest
        // of the icon. The catalog currently contains move-only SVG fragments.
      }
    }
    assert.equal(isFilledIcon(icon.id), icon.f === 1, `${icon.id} has inconsistent fill metadata`);
    const objects = icon.f === 1 ? fillSpans(runs).length : polylineSegments(runs).length;
    assert.ok(objects > 0, `${icon.id} produced no editable objects`);
  }
  assert.equal(index.icons.length, 54_250);
});

test("Make Editable converts a stroke icon into lines in its own box", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{
      id: "img1", type: "IMAGE", left: 100, top: 50, width: 48, height: 48,
      title: "IconAid: Frame", description: "Slide Aid icon. [iconaid:lucide-frame:#123456]",
    }],
    projectFiles: PROJECT_FILES,
  });

  const result = makeIconsEditable();
  assert.match(result.message, /Converted 1 icon/);
  assert.equal(state.elements.has("img1"), false, "the picture is replaced");

  const lines = requestsOfKind(state, "createLine");
  assert.equal(lines.length, 4, "a closed square is four segments");
  assert.equal(requestsOfKind(state, "createShape").length, 0, "a stroke icon draws no filled shapes");

  // The 24-unit design grid maps onto the picture's box: x=4 -> 100 + 4*(48/24).
  const first = lines[0]!.elementProperties as { transform: { translateX: number; translateY: number } };
  assert.equal(first.transform.translateX, 108);
  assert.equal(first.transform.translateY, 58);

  const stroke = requestsOfKind(state, "updateLineProperties")[0]! as {
    lineProperties: { lineFill: { solidFill: { color: { rgbColor: Record<string, number> } } }; weight: { magnitude: number } };
  };
  assert.equal(stroke.lineProperties.weight.magnitude, 3.2, "1.6 design stroke scaled by 48/24");
  const rgb = stroke.lineProperties.lineFill.solidFill.color.rgbColor;
  assert.ok(Math.abs(rgb.red - 0x12 / 255) < 1e-9, "the tagged colour is honoured");

  const group = requestsOfKind(state, "groupObjects")[0]!;
  assert.equal((group.childrenObjectIds as string[]).length, 4);
  const alt = requestsOfKind(state, "updatePageElementAltText").at(-1)!;
  assert.match(String(alt.description), /\[iconaid:lucide-frame:#123456\]/, "the tag survives conversion");
});

test("Make Editable fills a solid icon instead of outlining it", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{
      id: "img1", type: "IMAGE", left: 0, top: 0, width: 24, height: 24,
      title: "IconAid: Box", description: "[iconaid:bootstrap-box:#000000]",
    }],
    projectFiles: PROJECT_FILES,
  });

  makeIconsEditable();
  const shapes = requestsOfKind(state, "createShape");
  assert.ok(shapes.length > 20, "a solid glyph is scan-converted into slices");
  assert.equal(requestsOfKind(state, "createLine").length, 0, "no outline is drawn for a solid icon");
  assert.ok(shapes.every((shape) => shape.shapeType === "RECTANGLE"));
  // Slices are filled, not stroked.
  const fill = requestsOfKind(state, "updateShapeProperties")[0]! as { fields: string };
  assert.match(fill.fields, /shapeBackgroundFill\.solidFill/);
});

test("Make Editable reports clearly when there is nothing to convert", (t) => {
  t.after(teardown);
  installHost({ existing: [{ id: "s1", type: "SHAPE" }], projectFiles: PROJECT_FILES });
  assert.throws(() => makeIconsEditable(), /Select one or more inserted icons/);
});

test("Make Editable survives an icon missing from the catalog", (t) => {
  t.after(teardown);
  installHost({
    existing: [{ id: "img1", type: "IMAGE", description: "[iconaid:tabler-unknown:#000000]" }],
    projectFiles: PROJECT_FILES,
  });
  // Binary search finds no shard, so the icon is reported rather than crashing.
  assert.throws(() => makeIconsEditable(), /not in this deployment's catalog/);
});

function chartHost(cells: string[][]): HostState {
  return installHost({
    existing: [{ id: "table1", type: "TABLE", left: 20, top: 20, width: 200, height: 80, cells }],
    selectedIds: ["table1"],
  });
}

test("a column chart draws bars, value labels, a legend and an axis", (t) => {
  t.after(teardown);
  const state = chartHost([["", "2024", "2025"], ["Europe", "10", "20"], ["Asia", "5", "15"]]);

  buildChart("COL");

  assert.equal(state.batchUpdates, 1, "the whole chart is one atomic batch");
  const shapes = requestsOfKind(state, "createShape");
  const bars = shapes.filter((shape) => shape.shapeType === "RECTANGLE");
  // 4 data bars + 2 legend swatches.
  assert.equal(bars.length, 6);

  const text = insertedText(state);
  // Value labels for every datum - the thing that was missing entirely before.
  for (const value of ["10", "20", "5", "15"]) assert.ok(text.includes(value), `missing value label ${value}`);
  // Category labels and legend entries.
  for (const label of ["2024", "2025", "Europe", "Asia"]) assert.ok(text.includes(label), `missing label ${label}`);
  assert.ok(requestsOfKind(state, "createLine").length >= 1, "a baseline is drawn");

  // Each bar carries the datum it came from, for the annotation tools.
  const tags = requestsOfKind(state, "updatePageElementAltText").map((request) => String(request.description));
  assert.ok(tags.some((tag) => tag.includes(barTag(1, 1, 10))));
  assert.ok(tags.some((tag) => tag.includes(barTag(2, 2, 15))));
});

test("turning value labels and legend off removes them", (t) => {
  t.after(teardown);
  const state = chartHost([["", "2024"], ["Europe", "10"], ["Asia", "5"]]);

  applyChartSettings("COL", { ValueLabels: "0", Legend: "0", LabelSizePt: "9", Decimals: "auto", ClusterFill: "72" });
  state.requests.length = 0;
  buildChart("COL");

  const text = insertedText(state);
  assert.ok(!text.includes("10"), "value labels are suppressed");
  assert.ok(text.includes("2024"), "category labels remain");
  assert.ok(!text.includes("Europe"), "the legend is suppressed");
});

test("chart settings scope to the selected chart type and persist", (t) => {
  t.after(teardown);
  const state = installHost({ existing: [], selectedIds: [] });
  const before = chartSettingsState();
  assert.equal(before.scope, "GLOBAL", "nothing selected edits the new-chart defaults");
  assert.equal(before.values.ValueLabels, "1");

  applyChartSettings("COL", { ClusterFill: "50", ValueLabels: "0", Legend: "1", LabelSizePt: "12", Decimals: "1" });
  const stored = JSON.parse([...state.documentProperties.entries()].filter(([key]) => key.startsWith("slideAid.deck"))
    .sort()
    .filter(([key]) => !key.endsWith("chunks"))
    .map(([, value]) => value)
    .join("")) as { chartStyle: Record<string, string> };
  assert.equal(stored.chartStyle["COL.ClusterFill"], "0.5", "a percentage is stored as a ratio");
  assert.equal(stored.chartStyle["COL.ValueLabels"], "0");
  assert.equal(stored.chartStyle["COL.LabelSizePt"], "12");
  assert.equal(stored.chartStyle.ClusterFill, undefined, "a scoped edit never writes the global key");
});

test("family palettes are validated before they are stored", (t) => {
  t.after(teardown);
  installHost();
  assert.throws(() => saveFamilyPalette("BARS", ["#zzzzzz"]), /not a #RRGGBB colour|not a #RRGGBB color/);
  assert.throws(() => saveFamilyPalette("NOPE", ["#112233"]), /Unknown chart family/);
  assert.throws(() => saveFamilyPalette("BARS", []), /at least one colour|at least one color/);
  const saved = saveFamilyPalette("LINES", ["#112233", "#445566"]);
  assert.match(saved.message, /LINES palette saved \(2 colors\)/);
});

test("a Harvey ball is built from rotated slices and carries its state", (t) => {
  t.after(teardown);
  const state = installHost();
  insertHarveyBall(50, "#1F4E79");

  const shapes = requestsOfKind(state, "createShape");
  // Twelve of twenty-four slivers for half, plus the outline ring.
  assert.equal(shapes.length, 13);
  assert.equal(shapes.at(-1)!.shapeType, "ELLIPSE", "the ring is drawn last, on top");
  const outline = requestsOfKind(state, "updateShapeProperties").at(-1)! as { fields: string };
  assert.match(outline.fields, /outline/, "the ring is stroked, not filled");
  const alt = requestsOfKind(state, "updatePageElementAltText").at(-1)!;
  assert.match(String(alt.description), /\[slide-aid-harvey:50\]/, "Cycle State reads this back");

  // A full ball short-circuits to a single filled circle.
  state.requests.length = 0;
  insertHarveyBall(100, "#1F4E79");
  const full = requestsOfKind(state, "createShape");
  assert.equal(full.length, 2, "one filled disc plus the ring");
});

test("difference annotations need two tagged bars and use their data", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [
      { id: "chart1", type: "GROUP", description: "Chart Aid COL chart. [slide-aid-chart:abc123]" },
      { id: "bar1", type: "SHAPE", left: 100, top: 100, width: 20, height: 50, description: barTag(1, 1, 80) },
      { id: "bar2", type: "SHAPE", left: 200, top: 60, width: 20, height: 90, description: barTag(1, 2, 100) },
    ],
    selectedIds: ["chart1", "bar1", "bar2"],
  });
  // The chart's payload has to be readable for the annotation to resolve.
  const metadata = { schema: 1, id: "abc123", kind: "COL", data: { cells: [["", "a"], ["s", "1"]] }, rect: { left: 0, top: 0, width: 100, height: 100 }, palette: "Office", overrides: {} };
  const raw = JSON.stringify(metadata);
  state.documentProperties.set("slideAid.chart.v2.abc123.chunks", "1");
  state.documentProperties.set("slideAid.chart.v2.abc123.0", raw);

  const result = annotateDifference("PCT");
  assert.match(result.message, /\+25\.0%/, "percent difference from the stored data, not pixel heights");
  const text = insertedText(state);
  assert.ok(text.some((value) => value.includes("25.0")));
  assert.ok(requestsOfKind(state, "createLine").length >= 3, "a connector and two leaders");
});

test("annotations refuse a selection without two bars", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{ id: "chart1", type: "GROUP", description: "[slide-aid-chart:abc123]" }],
    selectedIds: ["chart1"],
  });
  state.documentProperties.set("slideAid.chart.v2.abc123.chunks", "1");
  state.documentProperties.set("slideAid.chart.v2.abc123.0", JSON.stringify({
    schema: 1, id: "abc123", kind: "COL", data: { cells: [["", "a"], ["s", "1"]] },
    rect: { left: 0, top: 0, width: 100, height: 100 }, palette: "Office", overrides: {},
  }));
  // The chart resolves, so the complaint is about the bars rather than the chart.
  assert.throws(() => annotateDifference("ABS"), /two bars/);
});

test("a chart with no payload is reported as not being a Chart Aid chart", (t) => {
  t.after(teardown);
  installHost({ existing: [{ id: "shape1", type: "SHAPE" }], selectedIds: ["shape1"] });
  assert.throws(() => annotateDifference("ABS"), /Select a Chart Aid chart/);
});

// --- Review markup ---------------------------------------------------------

test("a comment mark is a loud tagged note stamped with initials", (t) => {
  t.after(teardown);
  const state = installHost({ userEmail: "eero.siivola@example.com", pageWidth: 720 });
  const result = addReviewNote("NOTE", "tighten this headline");

  const shapes = requestsOfKind(state, "createShape");
  assert.equal(shapes.length, 2, "a filled note plus a text box");
  assert.equal(shapes[0]!.shapeType, "ROUND_RECTANGLE");
  const fill = requestsOfKind(state, "updateShapeProperties")[0]! as { shapeProperties: { shapeBackgroundFill: { solidFill: { color: { rgbColor: { red: number } } } } } };
  assert.ok(fill.shapeProperties.shapeBackgroundFill.solidFill.color.rgbColor.red === 1, "loud yellow fill (R=1)");
  const text = insertedText(state).join("\n");
  assert.match(text, /ES ·/, "initials derived from the account email");
  assert.match(text, /tighten this headline/);
  const alt = requestsOfKind(state, "updatePageElementAltText").at(-1)!;
  assert.match(String(alt.description), /\[slide-aid-review:NOTE\]/);
  assert.match(result.message, /comment mark/i);
});

test("stored initials override the account-derived ones", (t) => {
  t.after(teardown);
  const state = installHost({ userEmail: "someone.else@example.com" });
  setReviewInitials("XY");
  addReviewNote("TODO", "fix");
  assert.match(insertedText(state).join("\n"), /TODO · XY ·/, "TODO label plus the chosen initials");
});

test("a callout adds a leader line to the selected object", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{ id: "img1", type: "IMAGE", left: 300, top: 200, width: 100, height: 80 }],
    selectedIds: ["img1"],
    userEmail: "ada.lovelace@example.com",
  });
  addReviewCallout("is this the right chart?");
  assert.equal(requestsOfKind(state, "createLine").length, 1, "one leader line");
  const alt = requestsOfKind(state, "updatePageElementAltText").at(-1)!;
  assert.match(String(alt.description), /\[slide-aid-review:CALLOUT\]/);
});

test("remove markup sweeps every tagged shape across the deck", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [
      { id: "note1", type: "SHAPE", description: "Slide Aid NOTE [slide-aid-review:NOTE] a" },
      { id: "todo1", type: "GROUP", description: "[slide-aid-review:TODO] b" },
      { id: "keep1", type: "SHAPE", description: "a normal shape" },
    ],
  });
  const result = removeReviewMarkup();
  assert.equal(result.removed, 2, "both review marks, not the normal shape");
  assert.ok(state.elements.has("keep1"));
  assert.ok(!state.elements.has("note1") && !state.elements.has("todo1"));
});

test("initials derive from a name or email local-part, capped at three", () => {
  assert.equal(initialsFrom("eero.siivola"), "ES");
  assert.equal(initialsFrom("Ada King Lovelace"), "AKL");
  assert.equal(initialsFrom("a_b_c_d"), "ABC");
  assert.equal(initialsFrom(""), "");
});

// --- Icon catalog sharding -------------------------------------------------

test("an icon resolves to its shard by binary search over the boundary table", (t) => {
  t.after(teardown);
  installHost({ projectFiles: PROJECT_FILES });
  // One id from each shard, plus ids that fall outside every range.
  assert.deepEqual(iconPathsFor(["lucide-frame"]), { "lucide-frame": [STROKE_PATH] });
  assert.deepEqual(iconPathsFor(["bootstrap-box"]), { "bootstrap-box": [SOLID_PATH, SOLID_PATH] });
  assert.deepEqual(iconPathsFor(["aaa-before-everything"]), {});
  assert.deepEqual(iconPathsFor(["zzz-after-everything"]), {});
  // A mixed request spans both shards and drops the unknown one.
  assert.deepEqual(Object.keys(iconPathsFor(["bootstrap-box", "nope-missing", "lucide-frame"])).sort(),
    ["bootstrap-box", "lucide-frame"]);
});

test("icon path requests are bounded and sanitised", (t) => {
  t.after(teardown);
  installHost({ projectFiles: PROJECT_FILES });
  assert.throws(() => iconPathsFor(Array.from({ length: 501 }, () => "lucide-frame")), /Too many icons/);
  assert.throws(() => iconPathsFor("lucide-frame" as unknown as string[]), /by id/);
  // Ids that could not come from the catalog are ignored rather than looked up.
  assert.deepEqual(iconPathsFor(["../../etc/passwd", "Lucide-Frame"]), {});
});

test("a deployment without the catalog says so plainly", (t) => {
  t.after(teardown);
  installHost({ projectFiles: {} });
  assert.throws(() => iconPathsFor(["lucide-frame"]), /icon catalog is missing from this deployment/);
});

// --- View & Expert ---------------------------------------------------------

test("hiding and unhiding restores the original position", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{ id: "s1", type: "SHAPE", left: 120, top: 60, width: 40, height: 20 }],
    selectedIds: ["s1"],
  });

  hideObjects();
  const hidden = state.elements.get("s1")!;
  assert.ok(hidden.left > 1000, "the object is parked off-canvas");
  assert.equal(hidden.top, 60, "only the horizontal position moves");
  assert.match(hidden.description, /\[slide-aid-hidden:120:60\]/);

  unhideAll();
  const restored = state.elements.get("s1")!;
  assert.equal(restored.left, 120);
  assert.equal(restored.top, 60);
  assert.ok(!restored.description.includes("slide-aid-hidden"), "the marker is cleared");
  // Nothing left to restore.
  assert.throws(() => unhideAll(), /No hidden objects/);
});

test("hiding twice does not lose the original position", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{ id: "s1", type: "SHAPE", left: 10, top: 10, width: 5, height: 5 }],
    selectedIds: ["s1"],
  });
  hideObjects();
  // A second hide must be refused, or the parked coordinate would be recorded
  // as the "real" one and the object would never come back.
  assert.throws(() => hideObjects(), /already hidden/);
  unhideAll();
  assert.equal(state.elements.get("s1")!.left, 10);
});

// --- Snap to Table ---------------------------------------------------------

test("snap to table centres objects in the cell beneath them", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [
      { id: "table1", type: "TABLE", left: 0, top: 0, width: 200, height: 100, cells: [["", ""], ["", ""]] },
      // Sits over the second column, first row.
      { id: "dot", type: "SHAPE", left: 130, top: 20, width: 10, height: 10 },
    ],
    selectedIds: ["table1", "dot"],
  });
  const table = state.elements.get("table1")!;
  (table as unknown as { columnWidths: number[]; rowHeights: number[] }).columnWidths = [100, 100];
  (table as unknown as { columnWidths: number[]; rowHeights: number[] }).rowHeights = [50, 50];

  snapToTable("C");
  const dot = state.elements.get("dot")!;
  assert.equal(dot.left, 145, "centred in the 100-200 column");
  assert.equal(dot.top, 20, "centred in the 0-50 row");
});

test("snap to table needs both the table and something to snap", (t) => {
  t.after(teardown);
  installHost({ existing: [{ id: "a", type: "SHAPE" }, { id: "b", type: "SHAPE" }], selectedIds: ["a", "b"] });
  assert.throws(() => snapToTable("C"), /Select the table/);
});

// --- Geometry command dispatch ---------------------------------------------

test("Magic Resizer scales geometry and type together", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [{
      id: "s1", type: "SHAPE", left: 100, top: 100, width: 50, height: 20,
      runs: [{ text: "KPI", fontSize: 12 }],
    }],
    selectedIds: ["s1"],
  });
  executeCommand({ command: "scale", scalePercent: 200, scaleFonts: true });
  const shape = state.elements.get("s1")!;
  assert.equal(shape.width, 100);
  assert.equal(shape.height, 40);
  assert.equal(shape.left, 75, "scaled about its centre");
  assert.deepEqual(state.fontSizes.get("s1"), [24], "12pt text doubled with the box");

  executeCommand({ command: "scale", scalePercent: 50, scaleFonts: false });
  assert.deepEqual(state.fontSizes.get("s1"), [24], "type is left alone when asked");
});

test("one-click Matrix needs no column count", (t) => {
  t.after(teardown);
  const state = installHost({
    existing: [
      { id: "a", type: "SHAPE", left: 0, top: 0, width: 10, height: 10 },
      { id: "b", type: "SHAPE", left: 40, top: 0, width: 10, height: 10 },
      { id: "c", type: "SHAPE", left: 80, top: 0, width: 10, height: 10 },
      { id: "d", type: "SHAPE", left: 120, top: 0, width: 10, height: 10 },
    ],
  });
  executeCommand({ command: "matrix", columns: 0, gapCm: 0 });
  // Four objects make a 2x2 near-square grid.
  assert.deepEqual([...state.elements.values()].map((item) => [item.left, item.top]),
    [[0, 0], [10, 0], [0, 10], [10, 10]]);
});
