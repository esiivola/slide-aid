import test from "node:test";
import assert from "node:assert/strict";
import { installHost, insertedText, requestsOfKind, restoreHost, type HostState } from "./host";
import { insertIconImage, makeIconsEditable } from "../src/slides/icons";
import { buildChart, applyChartSettings, chartSettingsState, saveFamilyPalette } from "../src/charts/charts";
import { annotateDifference, insertHarveyBall } from "../src/charts/annotations";
import { barTag } from "../src/core/chart-data";

// A tiny two-shard catalog: one stroke icon and one solid one.
const STROKE_PATH = "M4 4 L20 4 L20 20 L4 20 Z";
const SOLID_PATH = "M2 2 L22 2 L22 22 L2 22 Z";
const PROJECT_FILES = {
  IconShards: JSON.stringify([
    { shard: 0, firstId: "bootstrap-box", lastId: "bootstrap-box" },
    { shard: 1, firstId: "lucide-frame", lastId: "lucide-frame" },
  ]),
  IconPaths00: JSON.stringify({ "bootstrap-box": [SOLID_PATH] }),
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
