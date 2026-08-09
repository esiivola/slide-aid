import type { ChartKind } from "../core/chart-data";
import { activeContext } from "../slides/selection";
import { buildChart } from "./charts";

// The 14 layouts from docs/CHARTS.md, in Chart Aid ribbon order. Data Layouts
// shows these as text; Sample Slides builds one real slide per kind from the
// same rows, so the examples can never drift from the builders.
interface SampleDef {
  kind: ChartKind;
  title: string;
  layout: string;
  cells: string[][];
}

const GRID: string[][] = [
  ["", "2024", "2025", "2026"],
  ["Europe", "120", "138", "151"],
  ["Americas", "95", "104", "126"],
  ["Asia", "60", "78", "99"],
];

export const SAMPLES: readonly SampleDef[] = [
  { kind: "COL", title: "Column", layout: "Row 1 = categories (top-left cell empty), column 1 = series names, body = numbers.", cells: GRID },
  { kind: "BAR", title: "Bar", layout: "Same as Column; the value axis runs horizontally.", cells: GRID },
  { kind: "STK", title: "Stacked", layout: "Same as Column; segments stack and totals print above each column.", cells: GRID },
  { kind: "SBR", title: "Stacked bar", layout: "Same as Stacked, rotated to horizontal bars.", cells: GRID },
  { kind: "PCT", title: "100%", layout: "Same as Stacked; every column is normalized to 100%.", cells: GRID },
  {
    kind: "WF", title: "Waterfall",
    layout: "Rows of label | value. A value of 'e' or '=' becomes a computed subtotal bar.",
    cells: [["Opening", "100"], ["Price", "18"], ["Volume", "-7"], ["Mix", "12"], ["Subtotal", "e"], ["FX", "-9"], ["Closing", "="]],
  },
  { kind: "MEK", title: "Mekko", layout: "Same as Column; column widths follow column totals and segments show shares.", cells: GRID },
  { kind: "LINE", title: "Line", layout: "Same as Column; one line with markers per series.", cells: GRID },
  { kind: "AREA", title: "Area", layout: "Same as Column; series stack as filled areas.", cells: GRID },
  {
    kind: "PIE", title: "Pie", layout: "Rows of label | value.",
    cells: [["Retail", "45"], ["Wholesale", "30"], ["Online", "25"]],
  },
  {
    kind: "DON", title: "Doughnut", layout: "Rows of label | value.",
    cells: [["Retail", "45"], ["Wholesale", "30"], ["Online", "25"]],
  },
  {
    kind: "SCAT", title: "Scatter", layout: "Rows of label | x | y.",
    cells: [["Alpha", "12", "34"], ["Beta", "25", "18"], ["Gamma", "38", "46"], ["Delta", "47", "29"]],
  },
  {
    kind: "BUB", title: "Bubble", layout: "Rows of label | x | y | size - a fourth column turns scatter into bubbles.",
    cells: [["Alpha", "12", "34", "8"], ["Beta", "25", "18", "20"], ["Gamma", "38", "46", "14"], ["Delta", "47", "29", "30"]],
  },
  {
    kind: "GANTT", title: "Gantt", layout: "Rows of activity | start | end. Numbers or dates like 1.3.2026; start = end draws a milestone diamond.",
    cells: [["Discovery", "1", "3"], ["Design", "3", "6"], ["Build", "5", "11"], ["Go live", "12", "12"]],
  },
];

/** Text for the Data Layouts panel - the equivalent of PowerPoint's ChHelp dialog. */
export function dataLayouts(): { kind: string; title: string; layout: string; example: string }[] {
  return SAMPLES.map((sample) => ({
    kind: sample.kind,
    title: sample.title,
    layout: sample.layout,
    example: sample.cells.slice(0, 3).map((row) => row.join(" | ")).join("\n"),
  }));
}

const SAMPLE_TAG = "[slide-aid-sample]";

/**
 * Appends one live example slide per chart type: a correctly formatted table and
 * the chart the real builder makes from it, so a table can be copied as a
 * starting point. Clean-up removes them again.
 */
export function insertSampleSlides(): { ok: true; message: string } {
  const context = activeContext();
  const presentation = context.presentation;
  for (const sample of SAMPLES) {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    slide.setSkipped(false);
    const heading = slide.insertTextBox(`${sample.title} — ${sample.layout}`, 24, 14, presentation.getPageWidth() - 48, 34);
    heading.getText().getTextStyle().setFontSize(11).setBold(true);
    const columns = Math.max(...sample.cells.map((row) => row.length));
    const table = slide.insertTable(sample.cells.length, columns, 24, 60, Math.max(180, columns * 62), Math.max(40, sample.cells.length * 20));
    sample.cells.forEach((row, r) => {
      for (let c = 0; c < columns; c += 1) table.getCell(r, c).getText().setText(row[c] ?? "");
    });
    table.setTitle(`${sample.title} sample data`);
    table.setDescription(SAMPLE_TAG);
    // Build through the real code path: select the table, then run the builder.
    table.select();
    buildChart(sample.kind);
    slide.getPageElements().forEach((element) => {
      if (element.getDescription().includes("[slide-aid-chart:")) element.setDescription(`${element.getDescription()} ${SAMPLE_TAG}`);
    });
  }
  return { ok: true, message: `Inserted ${SAMPLES.length} sample slides.` };
}

export function removeSampleSlides(): { ok: true; message: string } {
  const context = activeContext();
  const slides = context.presentation.getSlides();
  let removed = 0;
  for (const slide of slides) {
    const isSample = slide.getPageElements().some((element) => element.getDescription().includes(SAMPLE_TAG));
    if (!isSample) continue;
    slide.remove();
    removed += 1;
  }
  if (!removed) throw new Error("This presentation has no Chart Aid sample slides.");
  return { ok: true, message: `Removed ${removed} sample slide${removed === 1 ? "" : "s"}.` };
}
