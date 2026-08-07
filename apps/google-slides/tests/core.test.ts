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
  contrastRatio, decodeLibraryReference, encodeLibraryReference, extractGoogleFileId, isOutsideSlide, normalizeIconDefinition,
  normalizeLayout, projectLayout,
} from "../src/core/integrations";
import { insertSlideAidIcon } from "../src/entrypoints/index";

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

test("IconAid inserts and groups native Google Slides vectors", () => {
  const inserted: Array<Record<string, unknown>> = [];
  const byId = new Map<string, Record<string, unknown>>();
  let groupState: Record<string, unknown> | undefined;

  function register(kind: string, values: unknown[]): Record<string, unknown> {
    const id = `element-${inserted.length + 1}`;
    const state: Record<string, unknown> = { id, kind, values, removed: false };
    const element = {
      getObjectId: () => id,
      getLineFill: () => ({ setSolidFill: (color: string) => { state.lineColor = color; } }),
      setWeight: (weight: number) => { state.weight = weight; },
      getFill: () => ({
        setSolidFill: (color: string) => { state.fillColor = color; },
        setTransparent: () => { state.fillTransparent = true; },
      }),
      getBorder: () => ({
        setTransparent: () => { state.borderTransparent = true; },
        getLineFill: () => ({ setSolidFill: (color: string) => { state.borderColor = color; } }),
        setWeight: (weight: number) => { state.borderWeight = weight; },
      }),
      remove: () => { state.removed = true; },
    };
    Object.assign(state, element);
    inserted.push(state);
    byId.set(id, state);
    return state;
  }

  const slide = {
    insertLine: (...values: unknown[]) => register("line", values),
    insertShape: (...values: unknown[]) => register("shape", values),
    group: (elements: unknown[]) => {
      groupState = { elements, selected: false };
      return {
        setTitle: (title: string) => { groupState!.title = title; },
        setDescription: (description: string) => { groupState!.description = description; },
        select: () => { groupState!.selected = true; },
      };
    },
  };
  const presentation = {
    getPageWidth: () => 720,
    getPageHeight: () => 405,
    getPageElementById: (id: string) => byId.get(id),
    getSelection: () => ({
      getCurrentPage: () => ({
        getPageType: () => "SLIDE",
        asSlide: () => slide,
      }),
      getPageElementRange: () => null,
    }),
  };
  (globalThis as unknown as { SlidesApp: unknown }).SlidesApp = {
    getActivePresentation: () => presentation,
    PageType: { SLIDE: "SLIDE" },
    LineCategory: { STRAIGHT: "STRAIGHT" },
    ShapeType: { ELLIPSE: "ELLIPSE", RECTANGLE: "RECTANGLE" },
  };

  const result = insertSlideAidIcon({
    id: "host-smoke",
    name: "Host Smoke",
    category: "Technology",
    aliases: ["runtime test", "adapter check"],
    tags: ["host", "smoke", "vector", "native", "group", "style", "metadata", "insertion"],
    primitives: [
      { kind: "line", x1: 1, y1: 2, x2: 3, y2: 4 },
      { kind: "rect", x: 5, y: 6, width: 7, height: 8, filled: true },
      { kind: "ellipse", x: 14, y: 10, width: 6, height: 6, filled: false },
    ],
  }, "#123456");

  assert.equal(result.ok, true);
  assert.equal(result.message, "Inserted Host Smoke.");
  assert.deepEqual(inserted[0]!.values, ["STRAIGHT", 327, 172.5, 333, 178.5]);
  assert.deepEqual(inserted[1]!.values, ["RECTANGLE", 339, 184.5, 21, 24]);
  assert.equal(inserted[0]!.lineColor, "#123456");
  assert.equal(inserted[0]!.weight, 1.5);
  assert.equal(inserted[1]!.fillColor, "#123456");
  assert.equal(inserted[1]!.borderTransparent, true);
  assert.equal(inserted[2]!.fillTransparent, true);
  assert.equal(inserted[2]!.borderColor, "#123456");
  assert.equal(inserted[2]!.borderWeight, 1.5);
  assert.equal((groupState!.elements as unknown[]).length, 3);
  assert.equal(groupState!.title, "IconAid: Host Smoke");
  assert.equal(groupState!.description, "Editable IconAid vector icon [iconaid:host-smoke]");
  assert.equal(groupState!.selected, true);
});
