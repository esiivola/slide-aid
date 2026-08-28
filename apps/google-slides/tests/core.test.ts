import test from "node:test";
import assert from "node:assert/strict";
import {
  align, bounds, distribute, fillGap, matchSize, matrix, placeRegion, processChain,
  scaleAroundCenter, setSpacing, squareColumns, stack, swapPositions, type Box,
} from "../src/core/geometry";
import {
  barTag, cagr, decodeMetadata, encodeMetadata, isSubtotal, paletteColor, parseBarTag, parseNumber,
  percentDifference, validateChartData, type ChartMetadata,
} from "../src/core/chart-data";
import {
  applyControlValues, clearScope, controlsFor, controlValues, familyForKind, formatValue,
  isKnownStyleKey, STYLE_KEYS, styleColor, styleFlag, styleNumber, styleString,
} from "../src/core/chart-style";
import {
  contrastRatio, decodeLibraryReference, encodeLibraryReference, extractGoogleFileId, isOutsideSlide, normalizeIconDefinition,
  normalizeLayout, projectLayout,
} from "../src/core/integrations";
import { flattenIconElements, flattenPath, polylineSegments } from "../src/core/icon-path";
import { fillSpans, isFilledIcon } from "../src/core/icon-fill";
import catalog from "../../../shared/iconaid/catalog.json" with { type: "json" };
import chartStyleSpec from "../../../shared/specs/chart-style.json" with { type: "json" };
import palettesSpec from "../../../shared/specs/palettes.json" with { type: "json" };
import chartKindsSpec from "../../../shared/specs/chart-kinds.json" with { type: "json" };
import { PALETTES } from "../src/storage/preferences";
import { SAMPLES, dataLayouts } from "../src/charts/samples";

const boxes: Box[] = [
  { id: "a", left: 10, top: 20, width: 30, height: 10 },
  { id: "b", left: 70, top: 30, width: 20, height: 20 },
  { id: "c", left: 120, top: 10, width: 10, height: 30 },
];

test("bounds includes every object's far edge", () => {
  assert.deepEqual(bounds(boxes), { id: "selection-bounds", left: 10, top: 10, width: 120, height: 40 });
});

test("alignment preserves target sizes", () => {
  const result = align(boxes.slice(0, 2), boxes[2]!, "R");
  assert.equal(result[0]!.left, 100);
  assert.equal(result[1]!.left, 110);
  assert.equal(result[0]!.width, 30);
});

test("size matching keeps target centers", () => {
  const result = matchSize([boxes[0]!], boxes[1]!, "WH")[0]!;
  assert.deepEqual(result, { ...boxes[0], left: 15, top: 15, width: 20, height: 20 });
});

test("ratio width matching scales height by the same factor and keeps center", () => {
  const source: Box = { id: "img", left: 0, top: 0, width: 10, height: 20 };
  const reference: Box = { id: "m", left: 0, top: 0, width: 30, height: 5 };
  const result = matchSize([source], reference, "WR")[0]!;
  // factor 3: width 10->30, height 20->60; center (5,10) preserved.
  assert.deepEqual(result, { id: "img", left: -10, top: -20, width: 30, height: 60 });
});

test("ratio height matching scales width by the same factor and keeps center", () => {
  const source: Box = { id: "img", left: 0, top: 0, width: 20, height: 10 };
  const reference: Box = { id: "m", left: 0, top: 0, width: 5, height: 30 };
  const result = matchSize([source], reference, "HR")[0]!;
  // factor 3: height 10->30, width 20->60; center (10,5) preserved.
  assert.deepEqual(result, { id: "img", left: -20, top: -10, width: 60, height: 30 });
});

test("fill gap extends only when target is on the requested side", () => {
  const reference = { id: "r", left: 40, top: 0, width: 10, height: 10 };
  const rightTarget = { id: "t", left: 70, top: 0, width: 20, height: 10 };
  assert.deepEqual(fillGap([rightTarget], reference, "L")[0], { ...rightTarget, left: 50, width: 40 });
  assert.deepEqual(fillGap([rightTarget], reference, "R")[0], rightTarget);
});

test("spacing uses deterministic spatial order", () => {
  const result = setSpacing([boxes[2]!, boxes[0]!, boxes[1]!], "H", 5);
  assert.deepEqual(result.map((box) => [box.id, box.left]), [["a", 10], ["b", 45], ["c", 70]]);
});

test("stack aligns the secondary coordinate", () => {
  const result = stack([boxes[2]!, boxes[0]!], "V", -2);
  assert.deepEqual(result.map((box) => [box.id, box.left, box.top]), [["c", 120, 10], ["a", 120, 38]]);
});

test("distribution keeps the overall envelope", () => {
  const result = distribute(boxes, "H");
  assert.equal(result[0]!.left, 10);
  assert.equal(result[2]!.left + result[2]!.width, 130);
  assert.equal(result[1]!.left - (result[0]!.left + result[0]!.width), 30);
});

test("matrix uses the largest cell dimensions", () => {
  const result = matrix(boxes, 2, 5, 7);
  assert.deepEqual(result.map((box) => [box.id, box.left, box.top]), [["a", 10, 20], ["b", 45, 20], ["c", 10, 57]]);
});

test("scale preserves centers", () => {
  assert.deepEqual(scaleAroundCenter([boxes[0]!], 2)[0], { ...boxes[0], left: -5, top: 15, width: 60, height: 20 });
});

test("place-region presets match PowerPoint proportions", () => {
  const slide = { id: "slide", left: 0, top: 0, width: 300, height: 180 };
  assert.deepEqual(placeRegion(slide, "Q4", 10), { id: "region-Q4", left: 160, top: 100, width: 130, height: 70 });
});

test("chart numbers accept comma decimals, spaces and percent signs", () => {
  assert.equal(parseNumber(" 1 234,5% "), 1234.5);
  assert.equal(parseNumber("not a number"), 0);
  assert.equal(isSubtotal(" E "), true);
  assert.equal(isSubtotal("="), true);
});

test("chart metadata round-trips without separator escaping bugs", () => {
  const metadata: ChartMetadata = {
    schema: 1,
    id: "chart1",
    kind: "COL",
    data: { cells: [["", "A|B"], ["Series", "1,5"]] },
    rect: { left: 1, top: 2, width: 3, height: 4 },
    palette: "Office",
    overrides: { "1": "#123456" },
  };
  assert.deepEqual(decodeMetadata(encodeMetadata(metadata)), metadata);
  assert.equal(decodeMetadata("unrelated alt text"), null);
});

test("chart validation catches wrong layouts", () => {
  assert.throws(() => validateChartData("COL", { cells: [["only one cell"]] }), /header row/);
  assert.doesNotThrow(() => validateChartData("WF", { cells: [["Revenue", "100"], ["Total", "="]] }));
});

test("series overrides take precedence over palette colors", () => {
  assert.equal(paletteColor(["#111111", "#222222"], 0, { "1": "#abcdef" }), "#abcdef");
  assert.equal(paletteColor(["#111111", "#222222"], 3), "#222222");
});

test("Google file IDs are accepted from URLs or raw IDs", () => {
  const id = "1AbCdEfGhIjKlMnOpQrStUvWxYz";
  assert.equal(extractGoogleFileId(`https://docs.google.com/spreadsheets/d/${id}/edit#gid=0`), id);
  assert.equal(extractGoogleFileId(id), id);
  assert.throws(() => extractGoogleFileId("not a Google file"), /valid Google/);
});

test("shared library references round-trip in accessible descriptions", () => {
  const reference = { presentationId: "presentation_123", slideId: "slide_456" };
  const description = `Reusable KPI card. ${encodeLibraryReference(reference)}`;
  assert.deepEqual(decodeLibraryReference(description), reference);
  assert.equal(decodeLibraryReference("ordinary alt text"), null);
});

test("named layouts normalize and project across slide sizes", () => {
  const slots = normalizeLayout([boxes[0]!, boxes[1]!], 200, 100);
  assert.deepEqual(slots[0], { left: 0.05, top: 0.2, width: 0.15, height: 0.1 });
  const projected = projectLayout(slots, [boxes[0]!, boxes[1]!], 400, 200);
  assert.deepEqual(projected[0], { ...boxes[0], left: 20, top: 40, width: 60, height: 20 });
  assert.throws(() => projectLayout(slots, [boxes[0]!], 400, 200), /requires 2/);
});

test("deck QA detects off-slide bounds", () => {
  assert.equal(isOutsideSlide({ id: "inside", left: 0, top: 0, width: 100, height: 50 }, 100, 50), false);
  assert.equal(isOutsideSlide({ id: "outside", left: 95, top: 0, width: 10, height: 10 }, 100, 50), true);
});

test("contrast ratios follow WCAG luminance math", () => {
  assert.equal(contrastRatio("#000000", "#FFFFFF"), 21);
  assert.ok(contrastRatio("#777777", "#FFFFFF") < 4.5);
});

test("IconAid definitions normalize supported vector primitives", () => {
  const icon = normalizeIconDefinition({
    id: "sample-icon",
    name: "Sample",
    category: "Technology",
    aliases: ["example icon", "sample symbol"],
    tags: ["example"],
    primitives: [
      { kind: "line", x1: 1, y1: 2, x2: 3, y2: 4 },
      { kind: "ellipse", x: 5, y: 6, width: 7, height: 8, filled: true },
    ],
    elements: [
      { kind: "path", d: "M4 4 C8 2 16 2 20 4", filled: false },
      { kind: "polyline", points: [[4, 10], [8, 12], [12, 10]], closed: false, filled: false },
    ],
  });
  assert.deepEqual(icon.primitives[1], { kind: "ellipse", x: 5, y: 6, width: 7, height: 8, filled: true });
  assert.deepEqual(icon.elements[0], { kind: "path", d: "M4 4 C8 2 16 2 20 4", filled: false });
  assert.deepEqual(icon.elements[1], { kind: "polyline", points: [[4, 10], [8, 12], [12, 10]], closed: false, filled: false });
});

test("IconAid definitions reject unsafe or unsupported payloads", () => {
  assert.throws(() => normalizeIconDefinition({
    id: "bad",
    name: "Bad",
    category: "Technology",
    aliases: ["bad icon", "unsupported symbol"],
    tags: ["bad"],
    primitives: [{ kind: "path", d: "M0 0" }, { kind: "line", x1: 0, y1: 0, x2: 1, y2: 1 }],
  }), /unsupported primitive/);
  assert.throws(() => normalizeIconDefinition({
    id: "outside",
    name: "Outside",
    category: "Technology",
    aliases: ["outside icon", "invalid symbol"],
    tags: ["bad"],
    primitives: [
      { kind: "line", x1: -1, y1: 0, x2: 1, y2: 1 },
      { kind: "line", x1: 0, y1: 0, x2: 1, y2: 1 },
    ],
  }), /between 0 and 24/);
  assert.throws(() => normalizeIconDefinition({
    id: "bad-path",
    name: "Bad Path",
    category: "Technology",
    aliases: ["bad path", "invalid path"],
    tags: ["bad"],
    primitives: [{ kind: "line", x1: 0, y1: 0, x2: 1, y2: 1 }, { kind: "line", x1: 1, y1: 1, x2: 2, y2: 2 }],
    elements: [{ kind: "path", d: "M0 0 L25 1" }],
  }), /path coordinates/);
});


// --- Icon path flattening --------------------------------------------------

test("flattenPath turns lines, axis moves and curves into point runs", () => {
  const [run] = flattenPath("M2 4 L8 4 H12 V10");
  assert.deepEqual(run!.points, [[2, 4], [8, 4], [12, 4], [12, 10]]);
  assert.equal(run!.closed, false);

  const [curve] = flattenPath("M0 0 C4 0 8 4 8 8");
  // Sampled, not a single segment, and it must land exactly on the endpoint.
  assert.ok(curve!.points.length > 3);
  assert.deepEqual(curve!.points[0], [0, 0]);
  assert.deepEqual(curve!.points[curve!.points.length - 1], [8, 8]);
});

test("flattenPath closes subpaths and honours relative commands", () => {
  const [closed] = flattenPath("M2 2 L6 2 L6 6 Z");
  assert.equal(closed!.closed, true);
  assert.deepEqual(closed!.points[closed!.points.length - 1], [2, 2]);

  const [relative] = flattenPath("M2 2 l4 0 l0 4");
  assert.deepEqual(relative!.points, [[2, 2], [6, 2], [6, 6]]);

  const runs = flattenPath("M0 0 L2 0 M6 6 L8 6");
  assert.equal(runs.length, 2);
  assert.deepEqual(runs[1]!.points, [[6, 6], [8, 6]]);
});

test("flattenPath rejects unsupported and malformed paths", () => {
  assert.throws(() => flattenPath("M0 0 A5 5 0 0 1 10 10"), /unsupported command/);
  assert.throws(() => flattenPath("L"), /incomplete line/);
  assert.throws(() => flattenPath(""), /empty/);
});

test("icon elements split into native shapes and line runs", () => {
  const { shapes, polylines } = flattenIconElements([
    { kind: "rect", x: 2, y: 2, width: 4, height: 4, filled: true },
    { kind: "ellipse", x: 8, y: 8, width: 4, height: 4, filled: false },
    { kind: "line", x1: 0, y1: 0, x2: 4, y2: 4 },
    { kind: "polyline", points: [[1, 1], [5, 1], [5, 5]], closed: true, filled: false },
    { kind: "path", d: "M10 10 L14 10", filled: false },
  ]);
  assert.equal(shapes.length, 2);
  assert.deepEqual(shapes.map((shape) => shape.kind), ["rect", "ellipse"]);
  // A closed polyline gains the return segment; the path and line each add one.
  assert.equal(polylines.length, 3);
  assert.deepEqual(polylines[1]!.points[polylines[1]!.points.length - 1], [1, 1]);

  const segments = polylineSegments(polylines);
  assert.equal(segments.length, 5);
  assert.deepEqual(segments[0], { x1: 0, y1: 0, x2: 4, y2: 4 });
});

test("polylineSegments drops zero-length segments", () => {
  const segments = polylineSegments([{ points: [[1, 1], [1, 1], [4, 1]], closed: false }]);
  assert.deepEqual(segments, [{ x1: 1, y1: 1, x2: 4, y2: 1 }]);
});

test("every catalog icon flattens to something drawable", () => {
  for (const icon of catalog.icons) {
    const definition = normalizeIconDefinition(icon);
    const { shapes, polylines } = flattenIconElements(definition.elements);
    const drawn = shapes.length + polylineSegments(polylines).length;
    assert.ok(drawn > 0, `${definition.id} produced nothing to draw`);
    // The insert path is what the sidebar previews, so it must never fall back
    // to the coarser `primitives` list.
    assert.ok(drawn >= definition.primitives.length - 2, `${definition.id} lost detail versus its primitives`);
  }
});

// --- Chart style system ----------------------------------------------------

test("per-type chart settings override the global ones", () => {
  const store = { LabelSizePt: "9", "COL.LabelSizePt": "12", ClusterFill: "0.8" };
  assert.equal(styleNumber(store, "COL", "LabelSizePt"), 12);
  assert.equal(styleNumber(store, "BAR", "LabelSizePt"), 9);
  assert.equal(styleNumber(store, "COL", "ClusterFill"), 0.8);
  // Unset keys fall back to the documented PowerPoint defaults.
  assert.equal(styleNumber({}, "COL", "StackFill"), 0.65);
  assert.equal(styleString({}, null, "Decimals"), "auto");
});

test("chart flags and colors read the way PowerPoint stores them", () => {
  assert.equal(styleFlag({}, "COL", "ValueLabels"), true);
  assert.equal(styleFlag({ "COL.ValueLabels": "0" }, "COL", "ValueLabels"), false);
  assert.equal(styleFlag({ Legend: "0" }, "STK", "Legend"), false);
  assert.equal(styleColor({ WaterfallUp: "9BBB59" }, "WF", "WaterfallUp"), "#9BBB59");
  // "theme" means "follow the palette", so callers get null and choose.
  assert.equal(styleColor({}, "GANTT", "GanttBarColor"), null);
  assert.equal(styleColor({ "GANTT.GanttBarColor": "#ff0000" }, "GANTT", "GanttBarColor"), "#FF0000");
});

test("value formatting matches FmtNum", () => {
  assert.equal(formatValue(1234, "auto"), "1,234");
  assert.equal(formatValue(1234.5, "auto"), "1,234.5");
  assert.equal(formatValue(1234.5, "0"), "1,235");
  assert.equal(formatValue(0.25, "2"), "0.25");
  assert.equal(formatValue(-98765.4, "auto"), "-98,765.4");
});

test("chart families follow the PowerPoint palette groups", () => {
  assert.equal(familyForKind("COL"), "BARS");
  assert.equal(familyForKind("GANTT"), "BARS");
  assert.equal(familyForKind("LINE"), "LINES");
  assert.equal(familyForKind("BUB"), "LINES");
  assert.equal(familyForKind("DON"), "PIES");
});

test("chart settings submissions are scoped and validated", () => {
  const patch = applyControlValues("COL", { ClusterFill: "80", ValueLabels: "1", Legend: "0", LabelSizePt: "11", Decimals: "1" });
  assert.deepEqual(patch, {
    "COL.ClusterFill": "0.8", "COL.ValueLabels": "1", "COL.Legend": "0",
    "COL.LabelSizePt": "11", "COL.Decimals": "1",
  });
  assert.deepEqual(applyControlValues("GLOBAL", { LabelSizePt: "10" }), { LabelSizePt: "10" });
  assert.throws(() => applyControlValues("COL", { LabelSizePt: "99" }), /between 5 and 24/);
  assert.throws(() => applyControlValues("COL", { Decimals: "7" }), /one of auto, 0, 1, 2/);
  assert.throws(() => applyControlValues("WF", { WaterfallUp: "nope" }), /#RRGGBB/);
  assert.equal(applyControlValues("WF", { WaterfallUp: "" })["WF.WaterfallUp"], "theme");
});

test("resetting a scope leaves the other scopes alone", () => {
  const store = { LabelSizePt: "9", "COL.LabelSizePt": "12", "BAR.LabelSizePt": "14" };
  assert.deepEqual(clearScope(store, "COL"), { LabelSizePt: "9", "BAR.LabelSizePt": "14" });
  // Clearing GLOBAL keeps per-type overrides, matching Reset to Defaults.
  assert.deepEqual(clearScope(store, "GLOBAL"), { "COL.LabelSizePt": "12", "BAR.LabelSizePt": "14" });
});

test("the settings panel offers each chart type only its own parameters", () => {
  assert.deepEqual(controlsFor("MEK").map((control) => control.key), ["MekkoGapPt", "LabelSizePt", "Decimals"]);
  assert.deepEqual(controlsFor("AREA").map((control) => control.key), ["LabelSizePt"]);
  assert.ok(controlsFor("WF").some((control) => control.key === "WaterfallTotal"));
  assert.ok(controlsFor("GLOBAL").some((control) => control.key === "PlotWidthCm"));
  // A percentage is shown as 0-100 even though it is stored as a ratio.
  assert.equal(controlValues({ "COL.ClusterFill": "0.72" }, "COL").ClusterFill, "72");
  assert.equal(controlValues({}, "COL").Legend, "1");
});

test("every style key the panels expose is a known key", () => {
  for (const scope of ["GLOBAL", "COL", "BAR", "STK", "SBR", "PCT", "WF", "MEK", "LINE", "AREA", "PIE", "DON", "SCAT", "BUB", "GANTT"]) {
    for (const control of controlsFor(scope)) {
      assert.ok(isKnownStyleKey(control.key), `${scope}.${control.key} is not a documented style key`);
    }
  }
});

// --- Annotation maths ------------------------------------------------------

test("bar tags round-trip the datum they were drawn from", () => {
  assert.deepEqual(parseBarTag(`chart ${barTag(2, 3, -12.5)}`), { series: 2, category: 3, value: -12.5 });
  assert.equal(parseBarTag("no tag here"), null);
});

test("CAGR and percent difference come from the data, not the pixels", () => {
  assert.ok(Math.abs(cagr(100, 121, 2) - 10) < 1e-9);
  assert.ok(Math.abs(percentDifference(80, 100) - 25) < 1e-9);
  assert.ok(Math.abs(percentDifference(100, 80) + 20) < 1e-9);
  assert.throws(() => cagr(100, 121, 0), /at least one period/);
  assert.throws(() => cagr(-1, 121, 2), /two positive values/);
  assert.throws(() => percentDifference(0, 10), /non-zero/);
});

// --- Geometry added for parity ---------------------------------------------

test("swap rotates positions through the selection", () => {
  const three: Box[] = [
    { id: "a", left: 0, top: 0, width: 10, height: 10 },
    { id: "b", left: 40, top: 0, width: 20, height: 20 },
    { id: "c", left: 80, top: 0, width: 10, height: 10 },
  ];
  const result = swapPositions(three, "C");
  // a takes b's centre, b takes c's, c takes a's.
  assert.equal(result[0]!.left, 45);
  assert.equal(result[1]!.left, 75);
  assert.equal(result[2]!.left, 0);
  assert.equal(result[0]!.width, 10, "sizes stay put unless asked for");

  const sized = swapPositions(three, "TL", true);
  assert.equal(sized[0]!.width, 20);
  assert.equal(sized[0]!.left, 40);
  assert.throws(() => swapPositions([three[0]!], "C"), /at least two/);
});

test("one-click matrix picks a near-square grid", () => {
  assert.equal(squareColumns(4), 2);
  assert.equal(squareColumns(6), 3);
  assert.equal(squareColumns(9), 3);
  assert.equal(squareColumns(10), 4);
  assert.equal(squareColumns(1), 1);
});

test("process chain matches the reference band and closes the gaps", () => {
  const arrows: Box[] = [
    { id: "a", left: 0, top: 5, width: 30, height: 10 },
    { id: "b", left: 50, top: 20, width: 40, height: 30 },
  ];
  const reference: Box = { id: "r", left: 0, top: 12, width: 30, height: 24, rotation: 15 };
  const result = processChain(arrows, reference);
  assert.deepEqual(result.map((box) => box.left), [0, 30]);
  assert.deepEqual(result.map((box) => box.top), [12, 12]);
  assert.deepEqual(result.map((box) => box.height), [24, 24]);
  assert.deepEqual(result.map((box) => box.rotation), [15, 15]);
});

// --- Solid icon fill -------------------------------------------------------

test("solid icon sources are recognised", () => {
  assert.equal(isFilledIcon("bootstrap-alarm"), true);
  assert.equal(isFilledIcon("heroicons-bell-solid"), true);
  assert.equal(isFilledIcon("lucide-bell"), false);
  assert.equal(isFilledIcon("tabler-flask-2"), false);
});

test("a solid region scan-converts to spans covering it", () => {
  const square = flattenPath("M4 4 L20 4 L20 20 L4 20 Z");
  const spans = fillSpans(square, 8);
  assert.equal(spans.length, 8);
  for (const span of spans) {
    assert.ok(Math.abs(span.left - 4) < 1e-6);
    assert.ok(Math.abs(span.width - 16) < 1e-6);
  }
  // The slices span the shape's full height, with a little overlap so no
  // hairline shows between them.
  assert.ok(Math.abs(spans[0]!.top - 4) < 1e-6);
  assert.ok(spans[spans.length - 1]!.top + spans[spans.length - 1]!.height >= 20);
});

test("even-odd fill leaves holes open", () => {
  // Outer box with an inner box: the inner one must punch through, which is what
  // makes a Bootstrap "0 circle" read as a ring rather than a disc.
  const ring = [...flattenPath("M0 0 L24 0 L24 24 L0 24 Z"), ...flattenPath("M8 8 L16 8 L16 16 L8 16 Z")];
  const spans = fillSpans(ring, 12);
  const middle = spans.filter((span) => span.top < 12 && span.top + span.height > 12);
  assert.equal(middle.length, 2, "a scanline through the hole yields two spans");
  assert.ok(middle[0]!.left + middle[0]!.width <= 8 + 1e-6);
  assert.ok(middle[1]!.left >= 16 - 1e-6);
  // A scanline above the hole is one solid run.
  const top = spans.filter((span) => span.top < 2);
  assert.equal(top.length, 1);
  assert.ok(Math.abs(top[0]!.width - 24) < 1e-6);
});

test("filled catalog icons produce drawable spans", () => {
  // A real solid glyph, flattened and scan-converted the way Make Editable does.
  const glyph = flattenPath("M12 2 C6.5 2 2 6.5 2 12 C2 17.5 6.5 22 12 22 C17.5 22 22 17.5 22 12 C22 6.5 17.5 2 12 2 Z");
  const spans = fillSpans(glyph, 24);
  assert.ok(spans.length >= 20);
  // Widest span sits near the middle of a circle.
  const widest = spans.reduce((best, span) => (span.width > best.width ? span : best), spans[0]!);
  assert.ok(widest.top > 8 && widest.top < 16, "a circle is widest across its middle");
  assert.ok(widest.width > 18);
});

// --- Shared contract: shared/specs must match the implementation -----------
//
// These files are described as the reviewable contract between the two
// products. Nothing enforced that, so drift was silent and only showed up as
// two platforms drawing different charts.

test("chart style keys and defaults match shared/specs/chart-style.json", () => {
  const spec = chartStyleSpec as { keys: { key: string; default: string; description: string }[]; paletteFamilies: Record<string, string[]> };
  assert.deepEqual(
    STYLE_KEYS.map((entry) => ({ key: entry.key, default: entry.value, description: entry.description })),
    spec.keys,
    "src/core/chart-style.ts and shared/specs/chart-style.json disagree",
  );
  // Every kind belongs to exactly the family the spec assigns it.
  for (const family of Object.keys(spec.paletteFamilies)) {
    if (family === "comment") continue;
    for (const kind of spec.paletteFamilies[family]!) {
      assert.equal(familyForKind(kind), family, `${kind} should be in ${family}`);
    }
  }
});

test("palettes match shared/specs/palettes.json", () => {
  assert.deepEqual(PALETTES, palettesSpec, "deck palettes have drifted from the shared spec");
  for (const name of Object.keys(PALETTES)) {
    assert.equal(PALETTES[name]!.length, 6, `${name} must have six colors`);
    for (const color of PALETTES[name]!) assert.match(color, /^#[0-9A-F]{6}$/, `${name} has a malformed color`);
  }
});

test("chart kinds match shared/specs/chart-kinds.json", () => {
  const spec = chartKindsSpec as { chartKinds: string[] };
  // Every kind the spec lists must build, style and sample cleanly.
  const sampled = SAMPLES.map((sample) => sample.kind);
  assert.deepEqual([...sampled].sort(), [...spec.chartKinds].sort(), "Sample Slides and the shared spec disagree");
  for (const kind of spec.chartKinds) {
    assert.ok(controlsFor(kind).length > 0, `${kind} has no Chart Settings controls`);
    assert.ok(["BARS", "LINES", "PIES"].includes(familyForKind(kind)), `${kind} has no palette family`);
  }
});

test("every chart type documents a data layout with a usable example", () => {
  const layouts = dataLayouts();
  assert.equal(layouts.length, 14);
  for (const layout of layouts) {
    assert.ok(layout.layout.length > 20, `${layout.kind} needs a real layout description`);
    assert.ok(layout.example.includes("|"), `${layout.kind} needs example rows`);
  }
  // The examples must actually satisfy the validator they are teaching.
  for (const sample of SAMPLES) {
    assert.doesNotThrow(() => validateChartData(sample.kind, { cells: sample.cells }), `${sample.kind} sample does not validate`);
  }
});
