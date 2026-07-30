import test from "node:test";
import assert from "node:assert/strict";
import {
  align, bounds, distribute, fillGap, matchSize, matrix, placeRegion,
  scaleAroundCenter, setSpacing, stack, type Box,
} from "../src/core/geometry";
import {
  decodeMetadata, encodeMetadata, isSubtotal, paletteColor, parseNumber,
  validateChartData, type ChartMetadata,
} from "../src/core/chart-data";
import {
  contrastRatio, decodeLibraryReference, encodeLibraryReference, extractGoogleFileId, isOutsideSlide, normalizeLayout, projectLayout,
} from "../src/core/integrations";

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
