import {
  cagr, parseBarTag, percentDifference,
  type BarReference, type ChartAxis, type ChartMetadata,
} from "../core/chart-data";
import { formatValue, styleNumber, styleString } from "../core/chart-style";
import { activeContext, elementBox } from "../slides/selection";
import { getDeckSettings, loadChartMetadata } from "../storage/document-state";
import { ShapeBatch } from "../slides/shape-batch";

type PageElement = GoogleAppsScript.Slides.PageElement;

const ANNOTATION_GREY = "#404040";
const ANNOTATION_DARK = "#595959";
const HARVEY_TAG = /\[slide-aid-harvey:(\d+)\]/;
const CHECKBOX_TAG = /\[slide-aid-check:(checked|crossed|empty)\]/;

interface ChartTarget {
  metadata: ChartMetadata;
  element: PageElement;
  bars: { reference: BarReference; box: ReturnType<typeof elementBox> }[];
}

/**
 * Resolves what the user has selected inside a chart. Selecting whole bars is
 * the PowerPoint gesture ("click one bar, then Cmd-click a second"); in Slides
 * the same thing works by entering the group and shift-clicking, and each bar
 * carries the datum it was drawn from.
 */
function chartTarget(minimumBars = 0): ChartTarget {
  const context = activeContext(1);
  let metadata: ChartMetadata | null = null;
  let element: PageElement | null = null;
  const bars: { reference: BarReference; box: ReturnType<typeof elementBox> }[] = [];

  for (const selected of context.elements) {
    const own = loadChartMetadata(selected.getDescription());
    if (own) {
      metadata = own;
      element = selected;
      continue;
    }
    const reference = parseBarTag(selected.getDescription());
    const parent = selected.getParentGroup();
    const inherited = parent ? loadChartMetadata(parent.getDescription()) : null;
    if (inherited) {
      metadata = inherited;
      element = parent as unknown as PageElement;
    }
    if (reference) bars.push({ reference, box: elementBox(selected) });
  }
  if (!metadata || !element) throw new Error("Select a Chart Aid chart, or bars inside one.");
  if (bars.length < minimumBars) {
    throw new Error(minimumBars === 2
      ? "Select exactly two bars inside the chart: click one, then shift-click the second."
      : `Select at least ${minimumBars} bar${minimumBars === 1 ? "" : "s"} inside the chart.`);
  }
  // Deterministic spatial order, the same rule the geometry tools use.
  bars.sort((a, b) => (a.box.left - b.box.left) || (a.box.top - b.box.top));
  return { metadata, element, bars };
}

function annotationBatch(): { batch: ShapeBatch; context: ReturnType<typeof activeContext> } {
  const context = activeContext();
  return { batch: new ShapeBatch(context.presentation.getId(), context.slide.getObjectId()), context };
}

function labelSizeFor(metadata: ChartMetadata): number {
  return styleNumber(getDeckSettings().chartStyle, metadata.kind, "LabelSizePt");
}

function decimalsFor(metadata: ChartMetadata): string {
  return styleString(getDeckSettings().chartStyle, metadata.kind, "Decimals");
}

/**
 * Difference / % difference / CAGR arrow between two selected bars. The number
 * comes from the bars' stored data, never from their pixel heights - the same
 * guarantee the PowerPoint version makes.
 */
export function annotateDifference(mode: "ABS" | "PCT" | "CAGR", periods?: number): { ok: true; message: string } {
  const { metadata, bars } = chartTarget(2);
  if (bars.length !== 2) throw new Error("Select exactly two bars inside the chart.");
  const [first, second] = bars as [typeof bars[0], typeof bars[0]];
  const size = labelSizeFor(metadata);

  let text: string;
  if (mode === "ABS") {
    const delta = second.reference.value - first.reference.value;
    text = `${delta >= 0 ? "+" : "−"}${formatValue(Math.abs(delta), decimalsFor(metadata))}`;
  } else if (mode === "PCT") {
    const delta = percentDifference(first.reference.value, second.reference.value);
    text = `${delta >= 0 ? "+" : "−"}${formatValue(Math.abs(delta), "1")}%`;
  } else {
    const span = Number(periods);
    if (!Number.isFinite(span) || span < 1) throw new Error("Enter the number of periods for the CAGR (at least 1).");
    text = `CAGR ${formatValue(cagr(first.reference.value, second.reference.value, span), "1")}%`;
  }

  const { batch, context } = annotationBatch();
  const startX = first.box.left + first.box.width / 2;
  const endX = second.box.left + second.box.width / 2;
  const top = Math.min(first.box.top, second.box.top) - 26;
  batch.addLine(startX, top + 12, endX, top + 12, ANNOTATION_DARK, 1);
  batch.addLine(startX, top + 12, startX, first.box.top - 2, ANNOTATION_DARK, 0.75, true);
  batch.addLine(endX, top + 12, endX, second.box.top - 2, ANNOTATION_DARK, 0.75, true);
  batch.addText(text, Math.min(startX, endX), top - 4, Math.max(40, Math.abs(endX - startX)), size + 6, size, "CENTER", ANNOTATION_GREY, true);
  const id = batch.commit("Chart Aid annotation", `Chart Aid ${mode} annotation [slide-aid-annotation:${metadata.id}]`);
  const inserted = context.presentation.getPageElementById(id);
  if (inserted) inserted.select();
  return { ok: true, message: `Added ${text}.` };
}

/**
 * Chart kinds drawn against a single linear value axis. The others either have no
 * such axis (Mekko normalizes per column, scatter and bubble have two, Gantt is a
 * time band) or are rendered as an image, so a value line has nothing to sit on.
 */
const LINEAR_AXIS_KINDS = ["COL", "BAR", "STK", "SBR", "PCT", "WF", "LINE"];

function requireAxis(metadata: ChartMetadata): ChartAxis {
  if (metadata.axis) return metadata.axis;
  if (!LINEAR_AXIS_KINDS.includes(metadata.kind)) {
    throw new Error(`Value and average lines need a single value axis, which a ${metadata.kind} chart does not have.`);
  }
  // Charts built before the axis was recorded only need one rebuild.
  throw new Error("Rebuild this chart once so Slide Aid records its value scale, then add the line.");
}

function drawValueLine(metadata: ChartMetadata, value: number, caption: string): { ok: true; message: string } {
  const axis = requireAxis(metadata);
  const size = labelSizeFor(metadata);
  const { batch, context } = annotationBatch();
  if (axis.horizontal) {
    const x = axis.zero + value * axis.scale;
    batch.addLine(x, axis.plotStart, x, axis.plotStart + axis.plotSize, ANNOTATION_DARK, 1, true);
    batch.addText(caption, x - 30, axis.plotStart - size - 6, 60, size + 5, size, "CENTER", ANNOTATION_GREY, true);
  } else {
    const y = axis.zero - value * axis.scale;
    batch.addLine(axis.plotStart, y, axis.plotStart + axis.plotSize, y, ANNOTATION_DARK, 1, true);
    batch.addText(caption, axis.plotStart + axis.plotSize - 60, y - size - 4, 60, size + 5, size, "END", ANNOTATION_GREY, true);
  }
  const id = batch.commit("Chart Aid value line", `Chart Aid value line [slide-aid-annotation:${metadata.id}]`);
  const inserted = context.presentation.getPageElementById(id);
  if (inserted) inserted.select();
  return { ok: true, message: `Added a line at ${caption}.` };
}

export function annotateValueLine(value: number): { ok: true; message: string } {
  if (!Number.isFinite(value)) throw new Error("Enter the value the line should sit at.");
  const { metadata } = chartTarget();
  return drawValueLine(metadata, value, formatValue(value, decimalsFor(metadata)));
}

export function annotateAverageLine(): { ok: true; message: string } {
  const { metadata, bars } = chartTarget(2);
  const average = bars.reduce((sum, bar) => sum + bar.reference.value, 0) / bars.length;
  return drawValueLine(metadata, average, `Avg ${formatValue(average, decimalsFor(metadata))}`);
}

// ---------------------------------------------------------------------------
// Elements: Harvey balls and checkboxes.
//
// Google Slides exposes no adjustment handles and no freeform paths, so an
// arbitrary pie wedge cannot be made from a preset shape the way PowerPoint's
// msoShapePie can. The filled fraction is instead fanned out of the center as
// thin rotated slivers: a real vector, correct at any percentage, and close
// enough that the arc reads as round (a 15-degree sliver's flat outer edge sits
// under 1% of the radius inside the true arc).
// ---------------------------------------------------------------------------

const WEDGE_SLIVERS = 24;

function fillWedge(batch: ShapeBatch, cx: number, cy: number, radius: number, percent: number, color: string): void {
  if (percent <= 0) return;
  if (percent >= 100) {
    batch.addShape("ELLIPSE", cx - radius, cy - radius, radius * 2, radius * 2, color);
    return;
  }
  const step = (2 * Math.PI) / WEDGE_SLIVERS;
  const sweep = (percent / 100) * 2 * Math.PI;
  const count = Math.max(1, Math.round(sweep / step));
  // A sliver points straight down in its own coordinates, so half a turn puts
  // the first one at 12 o'clock; each following one advances clockwise.
  const width = 2 * radius * Math.tan(step / 2) + 0.3;
  for (let index = 0; index < count; index += 1) {
    batch.addPivotedShape("RECTANGLE", cx, cy, width, radius, Math.PI + index * step + step / 2, color);
  }
}

export function insertHarveyBall(percent: number, color: string, sizePt = 18): { ok: true; message: string } {
  const value = Math.round(Number(percent));
  if (!Number.isFinite(value) || value < 0 || value > 100) throw new Error("Enter a percentage between 0 and 100.");
  if (!/^#[0-9a-f]{6}$/i.test(color)) throw new Error("Choose a valid color.");
  const context = activeContext();
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  const left = (context.presentation.getPageWidth() - sizePt) / 2;
  const top = (context.presentation.getPageHeight() - sizePt) / 2;
  const radius = sizePt / 2;
  fillWedge(batch, left + radius, top + radius, radius, value, color);
  batch.addOutlinedShape("ELLIPSE", left, top, sizePt, sizePt, color, 1);
  const id = batch.commit(`Harvey ball ${value}%`, `Chart Aid Harvey ball [slide-aid-harvey:${value}]`);
  const inserted = context.presentation.getPageElementById(id);
  if (inserted) inserted.select();
  return { ok: true, message: `Inserted a ${value}% Harvey ball.` };
}

export function insertCheckbox(state: string, color: string, sizePt = 14): { ok: true; message: string } {
  const wanted = ["checked", "crossed", "empty"].includes(state) ? state : "checked";
  if (!/^#[0-9a-f]{6}$/i.test(color)) throw new Error("Choose a valid color.");
  const context = activeContext();
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  const left = (context.presentation.getPageWidth() - sizePt) / 2;
  const top = (context.presentation.getPageHeight() - sizePt) / 2;
  batch.addOutlinedShape("RECTANGLE", left, top, sizePt, sizePt, color, 1);
  if (wanted === "checked") {
    batch.addLine(left + sizePt * 0.22, top + sizePt * 0.55, left + sizePt * 0.44, top + sizePt * 0.78, color, 1.5);
    batch.addLine(left + sizePt * 0.44, top + sizePt * 0.78, left + sizePt * 0.8, top + sizePt * 0.24, color, 1.5);
  } else if (wanted === "crossed") {
    batch.addLine(left + sizePt * 0.24, top + sizePt * 0.24, left + sizePt * 0.76, top + sizePt * 0.76, color, 1.5);
    batch.addLine(left + sizePt * 0.76, top + sizePt * 0.24, left + sizePt * 0.24, top + sizePt * 0.76, color, 1.5);
  }
  const id = batch.commit(`Checkbox ${wanted}`, `Chart Aid checkbox [slide-aid-check:${wanted}]`);
  const inserted = context.presentation.getPageElementById(id);
  if (inserted) inserted.select();
  return { ok: true, message: `Inserted a ${wanted} checkbox.` };
}

const CHECK_CYCLE: Record<string, string> = { checked: "crossed", crossed: "empty", empty: "checked" };

/**
 * Steps every selected checkbox and Harvey ball to its next state, matching the
 * PowerPoint cycle: checkboxes checked -> crossed -> empty, Harvey balls in
 * 25% increments.
 */
export function cycleElementState(color: string): { ok: true; message: string } {
  const context = activeContext(1);
  let changed = 0;
  for (const element of context.elements) {
    const description = element.getDescription();
    const harvey = description.match(HARVEY_TAG);
    const check = description.match(CHECKBOX_TAG);
    if (!harvey && !check) continue;
    const box = elementBox(element);
    const fillColor = /^#[0-9a-f]{6}$/i.test(color) ? color : "#1F4E79";
    element.remove();
    const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
    if (harvey) {
      const next = (Math.round(Number(harvey[1]) / 25) * 25 + 25) % 125;
      const value = next > 100 ? 0 : next;
      const radius = Math.min(box.width, box.height) / 2;
      fillWedge(batch, box.left + box.width / 2, box.top + box.height / 2, radius, value, fillColor);
      batch.addOutlinedShape("ELLIPSE", box.left, box.top, box.width, box.height, fillColor, 1);
      batch.commit(`Harvey ball ${value}%`, `Chart Aid Harvey ball [slide-aid-harvey:${value}]`);
    } else {
      const next = CHECK_CYCLE[check![1]!] ?? "checked";
      const size = Math.min(box.width, box.height);
      batch.addOutlinedShape("RECTANGLE", box.left, box.top, box.width, box.height, fillColor, 1);
      if (next === "checked") {
        batch.addLine(box.left + size * 0.22, box.top + size * 0.55, box.left + size * 0.44, box.top + size * 0.78, fillColor, 1.5);
        batch.addLine(box.left + size * 0.44, box.top + size * 0.78, box.left + size * 0.8, box.top + size * 0.24, fillColor, 1.5);
      } else if (next === "crossed") {
        batch.addLine(box.left + size * 0.24, box.top + size * 0.24, box.left + size * 0.76, box.top + size * 0.76, fillColor, 1.5);
        batch.addLine(box.left + size * 0.76, box.top + size * 0.24, box.left + size * 0.24, box.top + size * 0.76, fillColor, 1.5);
      }
      batch.commit(`Checkbox ${next}`, `Chart Aid checkbox [slide-aid-check:${next}]`);
    }
    changed += 1;
  }
  if (!changed) throw new Error("Select one or more Slide Aid checkboxes or Harvey balls.");
  return { ok: true, message: `Cycled ${changed} element${changed === 1 ? "" : "s"}.` };
}
