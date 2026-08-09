import {
  barTag, DATASHEET_PREFIX, isSubtotal, paletteColor,
  parseNumber, validateChartData, type ChartAxis, type ChartData, type ChartKind, type ChartMetadata, type ChartRect, type SheetSource,
} from "../core/chart-data";
import {
  applyControlValues, clearScope, controlsFor, controlValues, familyForKind, formatValue,
  PALETTE_FAMILIES, scopeLabel, styleColor, styleFlag, styleNumber, styleString,
  type PaletteFamily, type StyleStore,
} from "../core/chart-style";
import { activeContext, elementBox } from "../slides/selection";
import { currentPalette, getSettings, PALETTES, updateSettings } from "../storage/preferences";
import { chartDescription, getDeckSettings, loadChartMetadata, saveChartMetadata, updateDeckSettings } from "../storage/document-state";
import { readSheetData, refreshSheetData } from "../integrations/sheets";
import { ShapeBatch } from "../slides/shape-batch";

type PageElement = GoogleAppsScript.Slides.PageElement;
type Slide = GoogleAppsScript.Slides.Slide;

const CM_TO_PT = 28.3464567;
const AXIS_GREY = "#595959";
const LABEL_GREY = "#404040";
const CONNECTOR_GREY = "#A0A0A0";

// Vertical room the builders reserve, matching the PowerPoint layout constants.
const CATEGORY_BAND = 16;
const LEGEND_BAND = 16;
const BAR_CATEGORY_GUTTER = 60;

interface BuildContext {
  slide: Slide;
  rect: ChartRect;
  batch: ShapeBatch;
  palette: readonly string[];
  metadata: ChartMetadata;
  style: StyleStore;
  /** Chart kind the style lookups are scoped to, so "COL.LabelSizePt" resolves. */
  kind: ChartKind;
  /** Filled in by the builders so the axis survives into the chart's metadata. */
  out: { axis?: ChartAxis };
}

function num(ctx: BuildContext, key: string): number {
  return styleNumber(ctx.style, ctx.kind, key);
}

function flag(ctx: BuildContext, key: string): boolean {
  return styleFlag(ctx.style, ctx.kind, key);
}

function labelSize(ctx: BuildContext): number {
  return num(ctx, "LabelSizePt");
}

function valueText(ctx: BuildContext, value: number): string {
  return formatValue(value, styleString(ctx.style, ctx.kind, "Decimals"));
}

/** Mirrors LabelColorOn(): dark text on light fills, white on dark ones. */
function labelColorOn(hex: string): string {
  const clean = hex.replace("#", "");
  const r = Number.parseInt(clean.slice(0, 2), 16);
  const g = Number.parseInt(clean.slice(2, 4), 16);
  const b = Number.parseInt(clean.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#323232" : "#FFFFFF";
}

function addRect(ctx: BuildContext, left: number, top: number, width: number, height: number, color: string): string {
  return ctx.batch.addShape("RECTANGLE", left, top, Math.max(0.5, width), Math.max(0.5, height), color);
}

/** Stamps a drawn datum so Difference / CAGR / Average can read it back. */
function tagDatum(ctx: BuildContext, id: string, series: number, category: number, value: number): void {
  ctx.batch.note(id, "Chart Aid data point", barTag(series, category, value));
}

function addEllipse(ctx: BuildContext, left: number, top: number, width: number, height: number, color: string): void {
  ctx.batch.addShape("ELLIPSE", left, top, Math.max(0.5, width), Math.max(0.5, height), color);
}

function addLine(ctx: BuildContext, x1: number, y1: number, x2: number, y2: number, color = AXIS_GREY, weight = 1, dashed = false): void {
  ctx.batch.addLine(x1, y1, x2, y2, color, weight, dashed);
}

type Alignment = "START" | "CENTER" | "END";

function addLabel(
  ctx: BuildContext, text: string, left: number, top: number, width: number,
  size = labelSize(ctx), alignment: Alignment = "CENTER", color = LABEL_GREY,
): void {
  ctx.batch.addText(text, left, top, Math.max(1, width), size + 5, size, alignment, color);
}

function numericRange(values: number[]): { min: number; max: number } {
  let min = Math.min(0, ...values);
  let max = Math.max(0, ...values);
  if (Math.abs(max - min) < 1e-9) max = min + 1;
  return { min, max };
}

function gridData(data: ChartData): { categories: string[]; series: { name: string; values: number[] }[] } {
  const header = data.cells[0] ?? [];
  const categories = header.slice(1);
  const series = data.cells.slice(1).map((row, index) => ({
    name: row[0] || `Series ${index + 1}`,
    values: categories.map((_, column) => parseNumber(row[column + 1] ?? "")),
  }));
  return { categories, series };
}

/** Legend strip along the top, laid out like the PowerPoint builders'. */
function drawLegend(ctx: BuildContext, left: number, top: number, series: { name: string }[]): void {
  let x = left;
  const size = labelSize(ctx);
  series.forEach((item, index) => {
    addRect(ctx, x, top + 3, 8, 8, paletteColor(ctx.palette, index, ctx.metadata.overrides));
    addLabel(ctx, item.name, x + 11, top, 60, size, "START");
    x += 12 + item.name.length * 5.5 + 12;
  });
}

function drawColumns(ctx: BuildContext, stacked: boolean, normalized: boolean, horizontal: boolean): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const categoryCount = categories.length;
  if (!categoryCount || !series.length) return;

  const showValues = flag(ctx, "ValueLabels");
  const showTotals = flag(ctx, "TotalLabels");
  const showLegend = flag(ctx, "Legend") && series.length > 1;
  const size = labelSize(ctx);

  const legendBand = showLegend ? LEGEND_BAND : 0;
  const plotTop = ctx.rect.top + legendBand;
  const plotHeight = ctx.rect.height - CATEGORY_BAND - legendBand;
  const plotLeft = ctx.rect.left + (horizontal ? BAR_CATEGORY_GUTTER : 0);
  const plotWidth = ctx.rect.width - (horizontal ? BAR_CATEGORY_GUTTER : 0);

  const values = series.flatMap((item) => item.values);
  const totals = categories.map((_, c) => series.reduce((sum, item) => sum + Math.abs(item.values[c] ?? 0), 0));
  const signedTotals = categories.flatMap((_, c) => {
    let positive = 0;
    let negative = 0;
    series.forEach((item) => {
      const value = item.values[c] ?? 0;
      if (value >= 0) positive += value;
      else negative += value;
    });
    return [positive, negative];
  });
  const range = normalized ? { min: 0, max: 100 } : numericRange(stacked ? signedTotals : values);
  const span = range.max - range.min;
  const scale = (horizontal ? plotWidth : plotHeight) / span;
  const zero = horizontal ? plotLeft + (0 - range.min) * scale : plotTop + range.max * scale;

  if (showLegend) drawLegend(ctx, plotLeft, ctx.rect.top, series);

  const slot = (horizontal ? plotHeight : plotWidth) / categoryCount;
  const fillShare = stacked ? num(ctx, "StackFill") : num(ctx, "ClusterFill");

  categories.forEach((category, c) => {
    let cumulativePositive = 0;
    let cumulativeNegative = 0;
    const columnTotal = normalized ? (totals[c] || 1) : 0;

    series.forEach((item, s) => {
      const raw = item.values[c] ?? 0;
      const value = normalized ? (Math.abs(raw) / columnTotal) * 100 : raw;
      const color = paletteColor(ctx.palette, s, ctx.metadata.overrides);
      let barLeft: number;
      let barTop: number;
      let barWidth: number;
      let barHeight: number;

      if (stacked) {
        const thickness = slot * fillShare;
        if (horizontal) {
          barTop = plotTop + c * slot + (slot - thickness) / 2;
          barHeight = thickness;
          if (value >= 0) {
            barLeft = zero + cumulativePositive * scale;
            barWidth = value * scale;
            cumulativePositive += value;
          } else {
            barLeft = zero + (cumulativeNegative + value) * scale;
            barWidth = -value * scale;
            cumulativeNegative += value;
          }
        } else {
          barLeft = plotLeft + c * slot + (slot - thickness) / 2;
          barWidth = thickness;
          if (value >= 0) {
            barTop = zero - (cumulativePositive + value) * scale;
            barHeight = value * scale;
            cumulativePositive += value;
          } else {
            barTop = zero - cumulativeNegative * scale;
            barHeight = -value * scale;
            cumulativeNegative += value;
          }
        }
      } else {
        const thickness = (slot * fillShare) / series.length;
        if (horizontal) {
          barTop = plotTop + c * slot + slot * 0.14 + s * thickness;
          barHeight = thickness * 0.92;
          barLeft = value >= 0 ? zero : zero + value * scale;
          barWidth = Math.abs(value) * scale;
        } else {
          barLeft = plotLeft + c * slot + slot * 0.14 + s * thickness;
          barWidth = thickness * 0.92;
          barTop = value >= 0 ? zero - value * scale : zero;
          barHeight = Math.abs(value) * scale;
        }
      }

      barWidth = Math.max(0.5, barWidth);
      barHeight = Math.max(0.5, barHeight);
      tagDatum(ctx, addRect(ctx, barLeft, barTop, barWidth, barHeight, color), s + 1, c + 1, raw);

      if (!showValues) return;
      const text = normalized ? `${Math.round(value)}%` : valueText(ctx, value);
      if (stacked) {
        // Only label a segment that can actually hold the text.
        const fits = horizontal ? barWidth > 20 : barHeight > size + 2;
        if (fits) addLabel(ctx, text, barLeft, barTop + barHeight / 2 - (size + 5) / 2, barWidth, size, "CENTER", labelColorOn(color));
      } else if (horizontal) {
        addLabel(ctx, text, barLeft + barWidth + 3, barTop + barHeight / 2 - (size + 5) / 2, 44, size, "START");
      } else {
        addLabel(ctx, text, barLeft - 10, value >= 0 ? barTop - (size + 5) : barTop + barHeight + 1, barWidth + 20, size, "CENTER");
      }
    });

    if (stacked && !normalized && showTotals) {
      const total = cumulativePositive;
      if (horizontal) addLabel(ctx, valueText(ctx, total), zero + total * scale + 3, plotTop + c * slot + slot / 2 - (size + 5) / 2, 44, size, "START");
      else addLabel(ctx, valueText(ctx, total), plotLeft + c * slot, zero - total * scale - (size + 5), slot, size, "CENTER");
    }

    if (horizontal) addLabel(ctx, category, ctx.rect.left, plotTop + c * slot + slot / 2 - (size + 5) / 2, BAR_CATEGORY_GUTTER - 4, size, "END");
    else addLabel(ctx, category, plotLeft + c * slot, plotTop + plotHeight + 2, slot, size, "CENTER");
  });

  if (horizontal) addLine(ctx, zero, plotTop, zero, plotTop + plotHeight);
  else addLine(ctx, plotLeft, zero, plotLeft + plotWidth, zero);

  ctx.out.axis = {
    zero, scale, horizontal,
    plotStart: horizontal ? plotTop : plotLeft,
    plotSize: horizontal ? plotHeight : plotWidth,
  };
}

function drawWaterfall(ctx: BuildContext): void {
  const rows = ctx.metadata.data.cells;
  const bars: { label: string; value: number; total: boolean; base: number }[] = [];
  let running = 0;
  for (const row of rows) {
    const total = isSubtotal(row[1] ?? "");
    const value = total ? running : parseNumber(row[1] ?? "");
    bars.push({ label: row[0] ?? "", value, total, base: total ? 0 : running });
    if (!total) running += value;
  }
  const size = labelSize(ctx);
  const showValues = flag(ctx, "ValueLabels");
  const up = styleColor(ctx.style, ctx.kind, "WaterfallUp") ?? "#9BBB59";
  const down = styleColor(ctx.style, ctx.kind, "WaterfallDown") ?? "#C0504D";
  const subtotal = styleColor(ctx.style, ctx.kind, "WaterfallTotal") ?? "#BFBFBF";
  const fillShare = num(ctx, "WaterfallFill");

  const extrema = bars.flatMap((item) => [item.base, item.total ? item.value : item.base + item.value]);
  const range = numericRange(extrema);
  const plotTop = ctx.rect.top + 8;
  const plotHeight = ctx.rect.height - CATEGORY_BAND - 8;
  const scale = plotHeight / (range.max - range.min);
  const zero = plotTop + range.max * scale;
  const slot = ctx.rect.width / Math.max(bars.length, 1);
  const inset = (1 - fillShare) / 2;

  addLine(ctx, ctx.rect.left, zero, ctx.rect.left + ctx.rect.width, zero);
  bars.forEach((item, index) => {
    const end = item.total ? item.value : item.base + item.value;
    const start = item.total ? 0 : item.base;
    const y1 = zero - start * scale;
    const y2 = zero - end * scale;
    const color = item.total ? subtotal : item.value >= 0 ? up : down;
    const barLeft = ctx.rect.left + index * slot + slot * inset;
    const barWidth = slot * fillShare;
    tagDatum(ctx, addRect(ctx, barLeft, Math.min(y1, y2), barWidth, Math.abs(y2 - y1), color), 1, index + 1, item.total ? item.value : item.value);
    if (showValues) {
      const text = valueText(ctx, item.total ? item.value : item.value);
      addLabel(ctx, text, barLeft - 10, Math.min(y1, y2) - (size + 5), barWidth + 20, size, "CENTER");
    }
    addLabel(ctx, item.label, ctx.rect.left + index * slot, plotTop + plotHeight + 2, slot, size, "CENTER");
    if (!item.total && index < bars.length - 1) {
      addLine(ctx, barLeft + barWidth, y2, ctx.rect.left + (index + 1) * slot + slot * inset, y2, CONNECTOR_GREY, 0.75);
    }
  });
  ctx.out.axis = { zero, scale, horizontal: false, plotStart: ctx.rect.left, plotSize: ctx.rect.width };
}

function drawMekko(ctx: BuildContext): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const size = labelSize(ctx);
  const showValues = flag(ctx, "ValueLabels");
  const showLegend = flag(ctx, "Legend") && series.length > 1;
  const gap = num(ctx, "MekkoGapPt");
  const totals = categories.map((_, c) => series.reduce((sum, item) => sum + Math.abs(item.values[c] ?? 0), 0));
  const grand = totals.reduce((sum, value) => sum + value, 0);
  if (!grand) return;

  const legendBand = showLegend ? LEGEND_BAND : 0;
  const totalBand = size + 6;
  const plotTop = ctx.rect.top + legendBand + totalBand;
  const plotHeight = ctx.rect.height - legendBand - totalBand - CATEGORY_BAND;
  if (showLegend) drawLegend(ctx, ctx.rect.left, ctx.rect.top, series);

  let x = ctx.rect.left;
  categories.forEach((category, c) => {
    const columnWidth = (totals[c]! / grand) * ctx.rect.width;
    const barWidth = Math.max(0.5, columnWidth - gap);
    let y = plotTop + plotHeight;
    addLabel(ctx, valueText(ctx, totals[c]!), x, plotTop - totalBand, columnWidth, size, "CENTER");
    series.forEach((item, s) => {
      const share = totals[c] ? (Math.abs(item.values[c] ?? 0) / totals[c]!) : 0;
      const height = share * plotHeight;
      y -= height;
      const color = paletteColor(ctx.palette, s, ctx.metadata.overrides);
      tagDatum(ctx, addRect(ctx, x + gap / 2, y, barWidth, height, color), s + 1, c + 1, item.values[c] ?? 0);
      if (showValues && height > size + 2) {
        addLabel(ctx, `${Math.round(share * 100)}%`, x + gap / 2, y + height / 2 - (size + 5) / 2, barWidth, size, "CENTER", labelColorOn(color));
      }
    });
    addLabel(ctx, category, x, plotTop + plotHeight + 2, columnWidth, size, "CENTER");
    x += columnWidth;
  });
}

function drawLineChart(ctx: BuildContext): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const size = labelSize(ctx);
  const showValues = flag(ctx, "ValueLabels");
  const showLegend = flag(ctx, "Legend") && series.length > 1;
  const marker = num(ctx, "MarkerSizePt");
  const values = series.flatMap((item) => item.values);
  const range = numericRange(values);

  const legendBand = showLegend ? LEGEND_BAND : 0;
  const plotLeft = ctx.rect.left + 10;
  const plotTop = ctx.rect.top + legendBand;
  const plotWidth = ctx.rect.width - 20;
  const plotHeight = ctx.rect.height - CATEGORY_BAND - legendBand;
  const scale = plotHeight / (range.max - range.min);
  const zero = plotTop + range.max * scale;
  if (showLegend) drawLegend(ctx, plotLeft, ctx.rect.top, series);

  addLine(ctx, plotLeft, zero, plotLeft + plotWidth, zero);
  const slot = plotWidth / Math.max(categories.length, 1);
  series.forEach((item, s) => {
    const color = paletteColor(ctx.palette, s, ctx.metadata.overrides);
    item.values.forEach((value, c) => {
      const x = plotLeft + c * slot + slot / 2;
      const y = zero - value * scale;
      if (c > 0) addLine(ctx, x - slot, zero - (item.values[c - 1] ?? 0) * scale, x, y, color, 2);
      if (marker > 0) tagDatum(ctx, ctx.batch.addShape("ELLIPSE", x - marker / 2, y - marker / 2, marker, marker, color), s + 1, c + 1, value);
      if (showValues) addLabel(ctx, valueText(ctx, value), x - slot / 2, y - (size + 5) - marker, slot, size, "CENTER");
    });
  });
  categories.forEach((category, c) => addLabel(ctx, category, plotLeft + c * slot, plotTop + plotHeight + 2, slot, size, "CENTER"));
  ctx.out.axis = { zero, scale, horizontal: false, plotStart: plotLeft, plotSize: plotWidth };
}

function drawScatter(ctx: BuildContext, bubble: boolean): void {
  const rows = ctx.metadata.data.cells;
  const size = labelSize(ctx);
  const points = rows.map((row) => ({ label: row[0] ?? "", x: parseNumber(row[1] ?? ""), y: parseNumber(row[2] ?? ""), size: parseNumber(row[3] ?? "") }));
  const xr = numericRange(points.map((point) => point.x));
  const yr = numericRange(points.map((point) => point.y));
  const maxSize = Math.max(1, ...points.map((point) => point.size));
  const left = ctx.rect.left + 30;
  const top = ctx.rect.top + 8;
  const width = ctx.rect.width - 40;
  const height = ctx.rect.height - CATEGORY_BAND - 8;
  addLine(ctx, left, top, left, top + height);
  addLine(ctx, left, top + height, left + width, top + height);
  points.forEach((point, index) => {
    const x = left + ((point.x - xr.min) / (xr.max - xr.min)) * width;
    const y = top + height - ((point.y - yr.min) / (yr.max - yr.min)) * height;
    const diameter = bubble ? 6 + 22 * Math.sqrt(Math.max(0, point.size) / maxSize) : 8;
    addEllipse(ctx, x - diameter / 2, y - diameter / 2, diameter, diameter, paletteColor(ctx.palette, index, ctx.metadata.overrides));
    addLabel(ctx, point.label, x + diameter / 2 + 2, y - (size + 5) / 2, 70, size, "START");
  });
}

function dateOrNumber(value: string): number {
  const dotted = value.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (dotted) return Date.UTC(Number(dotted[3]), Number(dotted[2]) - 1, Number(dotted[1])) / 86400000;
  const numeric = parseNumber(value);
  if (numeric || value.trim() === "0") return numeric;
  const date = Date.parse(value);
  return Number.isFinite(date) ? date / 86400000 : 0;
}

function drawGantt(ctx: BuildContext): void {
  const size = labelSize(ctx);
  const rows = ctx.metadata.data.cells.map((row) => ({ label: row[0] ?? "", start: dateOrNumber(row[1] ?? ""), end: dateOrNumber(row[2] ?? "") }));
  const min = Math.min(...rows.map((row) => row.start));
  const max = Math.max(...rows.map((row) => row.end));
  const span = Math.max(1, max - min);
  const labelWidth = Math.min(110, ctx.rect.width * 0.32);
  const left = ctx.rect.left + labelWidth;
  const width = ctx.rect.width - labelWidth - 4;
  const rowHeight = (ctx.rect.height - 8) / Math.max(rows.length, 1);
  // "theme" means follow the chart palette, exactly like the PowerPoint default.
  const barColor = styleColor(ctx.style, ctx.kind, "GanttBarColor") ?? paletteColor(ctx.palette, 0, ctx.metadata.overrides);
  rows.forEach((row, index) => {
    addLabel(ctx, row.label, ctx.rect.left, ctx.rect.top + index * rowHeight + rowHeight / 2 - (size + 5) / 2, labelWidth - 4, size, "END");
    const x = left + ((row.start - min) / span) * width;
    const end = left + ((row.end - min) / span) * width;
    if (Math.abs(row.end - row.start) < 1e-9) {
      ctx.batch.addShape("DIAMOND", x - 5, ctx.rect.top + index * rowHeight + rowHeight / 2 - 5, 10, 10, barColor);
    } else addRect(ctx, x, ctx.rect.top + index * rowHeight + rowHeight * 0.25, Math.max(1, end - x), rowHeight * 0.5, barColor);
  });
}

function staticChart(ctx: BuildContext, kind: "AREA" | "PIE" | "DON"): PageElement {
  const palette = ctx.palette.map((_, index) => paletteColor(ctx.palette, index, ctx.metadata.overrides));
  const pixelsWide = Math.max(240, Math.round(ctx.rect.width * 1.5));
  const pixelsHigh = Math.max(160, Math.round(ctx.rect.height * 1.5));
  const legend = flag(ctx, "Legend") ? Charts.Position.RIGHT : Charts.Position.NONE;
  if (kind === "AREA") {
    const { categories, series } = gridData(ctx.metadata.data);
    const builder = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Category");
    series.forEach((item) => builder.addColumn(Charts.ColumnType.NUMBER, item.name));
    categories.forEach((category, c) => builder.addRow([category, ...series.map((item) => item.values[c] ?? 0)]));
    const chart = Charts.newAreaChart().setDataTable(builder).setDimensions(pixelsWide, pixelsHigh)
      .setColors([...palette]).setStacked().setBackgroundColor("white").setLegendPosition(legend)
      .setOption("fontSize", labelSize(ctx)).build();
    return ctx.slide.insertImage(chart.getBlob(), ctx.rect.left, ctx.rect.top, ctx.rect.width, ctx.rect.height) as unknown as PageElement;
  }
  const builder = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Label").addColumn(Charts.ColumnType.NUMBER, "Value");
  ctx.metadata.data.cells.forEach((row) => builder.addRow([row[0] ?? "", Math.abs(parseNumber(row[1] ?? ""))]));
  const pie = Charts.newPieChart().setDataTable(builder).setDimensions(pixelsWide, pixelsHigh).setColors([...palette])
    .setBackgroundColor("white").setLegendPosition(legend).setOption("fontSize", labelSize(ctx));
  if (kind === "DON") pie.setOption("pieHole", 0.45);
  return ctx.slide.insertImage(pie.build().getBlob(), ctx.rect.left, ctx.rect.top, ctx.rect.width, ctx.rect.height) as unknown as PageElement;
}

/**
 * Colors for one chart: an Edit Colors family palette when the deck has one,
 * otherwise the named Color Theme. Matches PowerPoint, where a theme writes all
 * three family palettes and the builders read their own family back.
 */
function resolvePalette(kind: ChartKind, paletteName: string): readonly string[] {
  const family = familyForKind(kind);
  const custom = getDeckSettings().familyPalettes[family];
  if (custom?.length) return custom;
  return PALETTES[paletteName] ?? currentPalette();
}

function tagChart(element: PageElement, metadata: ChartMetadata): void {
  saveChartMetadata(metadata);
  element.setTitle(`Slide Aid ${metadata.kind} chart`);
  element.setDescription(chartDescription(metadata));
  element.select();
}

function drawChart(slide: Slide, metadata: ChartMetadata): void {
  validateChartData(metadata.kind, metadata.data);
  const palette = resolvePalette(metadata.kind, metadata.palette);
  const style = getDeckSettings().chartStyle;
  const presentation = SlidesApp.getActivePresentation();
  const ctx: BuildContext = {
    slide, rect: metadata.rect, palette, metadata, style, kind: metadata.kind, out: {},
    batch: new ShapeBatch(presentation.getId(), slide.getObjectId()),
  };
  if (["AREA", "PIE", "DON"].includes(metadata.kind)) {
    tagChart(staticChart(ctx, metadata.kind as "AREA" | "PIE" | "DON"), metadata);
    return;
  }
  if (metadata.kind === "COL") drawColumns(ctx, false, false, false);
  if (metadata.kind === "BAR") drawColumns(ctx, false, false, true);
  if (metadata.kind === "STK") drawColumns(ctx, true, false, false);
  if (metadata.kind === "SBR") drawColumns(ctx, true, false, true);
  if (metadata.kind === "PCT") drawColumns(ctx, true, true, false);
  if (metadata.kind === "WF") drawWaterfall(ctx);
  if (metadata.kind === "MEK") drawMekko(ctx);
  if (metadata.kind === "LINE") drawLineChart(ctx);
  if (metadata.kind === "SCAT") drawScatter(ctx, false);
  if (metadata.kind === "BUB") drawScatter(ctx, true);
  if (metadata.kind === "GANTT") drawGantt(ctx);
  const id = ctx.batch.commit(`Slide Aid ${metadata.kind} chart`, chartDescription(metadata));
  saveChartMetadata(ctx.out.axis ? { ...metadata, axis: ctx.out.axis } : metadata);
  const element = SlidesApp.openById(presentation.getId()).getPageElementById(id);
  if (element) element.select();
}

function tableData(table: GoogleAppsScript.Slides.Table): ChartData {
  const cells: string[][] = [];
  for (let row = 0; row < table.getNumRows(); row += 1) {
    const values: string[] = [];
    for (let column = 0; column < table.getNumColumns(); column += 1) values.push(table.getCell(row, column).getText().asString().trim());
    cells.push(values);
  }
  return { cells };
}

function selectedInputs(): { context: ReturnType<typeof activeContext>; tableElement: PageElement | null; chartElement: PageElement | null; metadata: ChartMetadata | null } {
  const context = activeContext(1);
  let tableElement: PageElement | null = null;
  let chartElement: PageElement | null = null;
  let metadata: ChartMetadata | null = null;
  for (const selected of context.elements) {
    if (selected.getPageElementType() === SlidesApp.PageElementType.TABLE) tableElement = selected;
    const own = loadChartMetadata(selected.getDescription());
    if (own) {
      chartElement = selected;
      metadata = own;
      continue;
    }
    const parent = selected.getParentGroup();
    const inherited = parent ? loadChartMetadata(parent.getDescription()) : null;
    if (inherited) {
      chartElement = parent as unknown as PageElement;
      metadata = inherited;
    }
  }
  if (chartElement && metadata) {
    const legacyDescription = chartElement.getDescription().startsWith("SLIDE_AID_CHART_V1:");
    const duplicates = context.presentation.getSlides().flatMap((slide) => slide.getPageElements())
      .filter((element) => element.getDescription().includes(`[slide-aid-chart:${metadata!.id}]`));
    if (duplicates.length > 1) {
      metadata = { ...metadata, id: Utilities.getUuid().replace(/-/g, "").slice(0, 16) };
      saveChartMetadata(metadata);
      chartElement.setDescription(chartDescription(metadata));
    } else if (legacyDescription) chartElement.setDescription(chartDescription(metadata));
  }
  return { context, tableElement, chartElement, metadata };
}

export function selectedChartState(): { kind: ChartKind; source?: SheetSource } | null {
  const { metadata } = selectedInputs();
  return metadata ? { kind: metadata.kind, source: metadata.source } : null;
}

function defaultRect(context: ReturnType<typeof activeContext>, tableElement: PageElement | null): ChartRect {
  const style = getDeckSettings().chartStyle;
  const width = styleNumber(style, null, "PlotWidthCm") * CM_TO_PT;
  const height = styleNumber(style, null, "PlotHeightCm") * CM_TO_PT;
  if (tableElement) {
    const box = elementBox(tableElement);
    const left = Math.min(context.presentation.getPageWidth() - width - 10, box.left + box.width + 20);
    return { left: Math.max(10, left), top: box.top, width, height };
  }
  return {
    left: (context.presentation.getPageWidth() - width) / 2,
    top: (context.presentation.getPageHeight() - height) / 2,
    width,
    height,
  };
}

function buildChartWithData(kindText: string, suppliedData?: ChartData, suppliedSource?: SheetSource): { ok: true; message: string } {
  const kind = kindText as ChartKind;
  const { context, tableElement, chartElement, metadata } = selectedInputs();
  const data = suppliedData ?? (tableElement ? tableData(tableElement.asTable()) : metadata?.data);
  if (!data) throw new Error("Select a data table or an existing Slide Aid chart.");
  validateChartData(kind, data);
  const oldBox = chartElement ? elementBox(chartElement) : null;
  const rect = oldBox ? { left: oldBox.left, top: oldBox.top, width: oldBox.width, height: oldBox.height } : defaultRect(context, tableElement);
  const next: ChartMetadata = {
    schema: 1,
    id: metadata?.id ?? Utilities.getUuid().replace(/-/g, "").slice(0, 16),
    kind,
    data,
    rect,
    palette: getDeckSettings().palette ?? getSettings().palette,
    overrides: metadata?.overrides ?? {},
    source: suppliedSource ?? (tableElement ? undefined : metadata?.source),
  };
  drawChart(context.slide, next);
  if (chartElement) chartElement.remove();
  if (tableElement?.getDescription().startsWith(DATASHEET_PREFIX)) tableElement.remove();
  return { ok: true, message: `Built ${kind} chart.` };
}

export function buildChart(kindText: string): { ok: true; message: string } {
  return buildChartWithData(kindText);
}

export function buildLinkedChart(kindText: string, spreadsheetUrl: string, sheetName: string, rangeA1: string): { ok: true; message: string } {
  const { data, source } = readSheetData(spreadsheetUrl, sheetName, rangeA1);
  return buildChartWithData(kindText, data, source);
}

export function validateLinkedChart(kindText: string, spreadsheetUrl: string, sheetName: string, rangeA1: string): { message: string; rows: number; columns: number; preview: string[][] } {
  const kind = kindText as ChartKind;
  const { data } = readSheetData(spreadsheetUrl, sheetName, rangeA1);
  validateChartData(kind, data);
  const columns = Math.max(...data.cells.map((row) => row.length));
  return { message: `Valid ${kind} source: ${data.cells.length} rows × ${columns} columns.`, rows: data.cells.length, columns, preview: data.cells.slice(0, 4).map((row) => row.slice(0, 5)) };
}

export function rebuildChart(): { ok: true; message: string } {
  const { metadata } = selectedInputs();
  if (!metadata) throw new Error("Select a Slide Aid chart.");
  if (metadata.source) return buildChartWithData(metadata.kind, refreshSheetData(metadata.source), metadata.source);
  return buildChartWithData(metadata.kind);
}

export function refreshLinkedChart(): { ok: true; message: string } {
  const { metadata } = selectedInputs();
  if (!metadata?.source) throw new Error("Select a chart linked to Google Sheets.");
  const result = buildChartWithData(metadata.kind, refreshSheetData(metadata.source), metadata.source);
  return { ...result, message: `Refreshed ${metadata.source.sheetName}!${metadata.source.rangeA1}.` };
}

export function editChartData(): { ok: true; message: string } {
  const { context, chartElement, metadata } = selectedInputs();
  if (!chartElement || !metadata) throw new Error("Select a Slide Aid chart.");
  const rows = metadata.data.cells.length;
  const columns = Math.max(...metadata.data.cells.map((row) => row.length));
  const box = elementBox(chartElement);
  const width = Math.max(160, columns * 60);
  const left = Math.max(5, box.left - width - 15);
  const table = context.slide.insertTable(rows, columns, left, box.top, width, Math.max(30, rows * 20));
  metadata.data.cells.forEach((row, r) => row.forEach((value, c) => table.getCell(r, c).getText().setText(value)));
  table.setTitle("Slide Aid chart datasheet");
  table.setDescription(DATASHEET_PREFIX + metadata.id);
  table.select();
  return { ok: true, message: "Edit the table, select it together with the chart, then click Rebuild." };
}

function allChartRecords(): { slide: Slide; element: PageElement; metadata: ChartMetadata }[] {
  const context = activeContext();
  const records: { slide: Slide; element: PageElement; metadata: ChartMetadata }[] = [];
  context.presentation.getSlides().forEach((slide) => slide.getPageElements().forEach((element) => {
    const metadata = loadChartMetadata(element.getDescription());
    if (metadata) records.push({ slide, element, metadata });
  }));
  return records;
}

export function restyleCharts(allCharts: boolean): { ok: true; message: string } {
  if (!allCharts) return rebuildChart();
  const records = allChartRecords();
  records.forEach(({ slide, element, metadata }) => {
    const box = elementBox(element);
    const next = { ...metadata, palette: getDeckSettings().palette ?? getSettings().palette, rect: { left: box.left, top: box.top, width: box.width, height: box.height } };
    drawChart(slide, next);
    element.remove();
  });
  return { ok: true, message: `Restyled ${records.length} chart${records.length === 1 ? "" : "s"}.` };
}

export function countCharts(): number {
  return allChartRecords().length;
}

export function setPalette(name: string): { ok: true; message: string; charts: number } {
  if (!PALETTES[name]) throw new Error(`Unknown palette: ${name}`);
  updateSettings({ palette: name });
  // A Color Theme is global in PowerPoint: it writes all three family palettes,
  // so any earlier Edit Colors work is intentionally replaced.
  const familyPalettes: Partial<Record<PaletteFamily, string[]>> = {};
  for (const family of PALETTE_FAMILIES) familyPalettes[family] = [...PALETTES[name]!];
  updateDeckSettings({ palette: name, familyPalettes });
  const charts = countCharts();
  return { ok: true, message: `Color theme set to ${name}.`, charts };
}

export function recolorSeries(seriesIndex: number, color: string): { ok: true; message: string } {
  if (!Number.isInteger(seriesIndex) || seriesIndex < 1) throw new Error("Series number must be a positive integer.");
  if (!/^#[0-9a-f]{6}$/i.test(color)) throw new Error("Choose a valid color.");
  const { context, chartElement, metadata } = selectedInputs();
  if (!chartElement || !metadata) throw new Error("Select a Slide Aid chart.");
  const box = elementBox(chartElement);
  const next: ChartMetadata = {
    ...metadata,
    rect: { left: box.left, top: box.top, width: box.width, height: box.height },
    overrides: { ...metadata.overrides, [String(seriesIndex)]: color },
  };
  drawChart(context.slide, next);
  chartElement.remove();
  return { ok: true, message: `Recolored series ${seriesIndex}.` };
}

// ---------------------------------------------------------------------------
// Chart Settings - the sidebar equivalent of the native PowerPoint panel.
// Scoped to the selected chart's kind, or to the new-chart defaults when nothing
// is selected, exactly like ChartSettingsDialog().
// ---------------------------------------------------------------------------

export interface ChartSettingsState {
  scope: string;
  scopeLabel: string;
  controls: { type: string; key: string; label: string; min?: number; max?: number; options?: string[] }[];
  values: Record<string, string>;
  hasSelectedChart: boolean;
}

export function chartSettingsState(): ChartSettingsState {
  let scope = "GLOBAL";
  let hasSelectedChart = false;
  try {
    const { metadata } = selectedInputs();
    if (metadata) {
      scope = metadata.kind;
      hasSelectedChart = true;
    }
  } catch {
    // Nothing selected: fall through to the global defaults, like PowerPoint.
  }
  return {
    scope,
    scopeLabel: scopeLabel(scope),
    controls: controlsFor(scope).map((control) => ({ ...control })),
    values: controlValues(getDeckSettings().chartStyle, scope),
    hasSelectedChart,
  };
}

export function applyChartSettings(scope: string, values: Record<string, unknown>): { ok: true; message: string; charts: number } {
  const patch = applyControlValues(scope, values);
  const chartStyle = { ...getDeckSettings().chartStyle, ...patch };
  updateDeckSettings({ chartStyle });
  if (scope === "GLOBAL") return { ok: true, message: "Defaults saved (used by new charts).", charts: countCharts() };
  // The settings are already stored, so losing the selection between opening the
  // panel and applying must not read as a failure - fall back to offering the
  // deck-wide restyle instead.
  try {
    const { metadata } = selectedInputs();
    if (metadata?.kind === scope) {
      rebuildChart();
      return { ok: true, message: `${scopeLabel(scope)} settings applied.`, charts: 0 };
    }
  } catch {
    // No chart selected any more; fall through.
  }
  return { ok: true, message: `${scopeLabel(scope)} settings saved.`, charts: countCharts() };
}

export function resetChartSettings(scope: string): { ok: true; message: string } {
  const current = getDeckSettings().chartStyle;
  updateDeckSettings({ chartStyle: scope === "ALL" ? {} : clearScope(current, scope) });
  return { ok: true, message: scope === "ALL" ? "All chart settings reset to defaults." : `${scopeLabel(scope)} settings reset to the defaults.` };
}

// ---------------------------------------------------------------------------
// Edit Colors - per-family palette editing (Bars / Lines / Pies).
// ---------------------------------------------------------------------------

export function familyPalette(family: string): { family: string; label: string; colors: string[] } {
  if (!PALETTE_FAMILIES.includes(family as PaletteFamily)) throw new Error(`Unknown chart family: ${family}`);
  const deck = getDeckSettings();
  const stored = deck.familyPalettes[family as PaletteFamily];
  const colors = stored?.length ? [...stored] : [...(PALETTES[deck.palette ?? getSettings().palette] ?? currentPalette())];
  return { family, label: family, colors };
}

export function saveFamilyPalette(family: string, colors: unknown): { ok: true; message: string; charts: number } {
  if (!PALETTE_FAMILIES.includes(family as PaletteFamily)) throw new Error(`Unknown chart family: ${family}`);
  if (!Array.isArray(colors) || !colors.length) throw new Error("A palette needs at least one color.");
  if (colors.length > 24) throw new Error("A palette can hold at most 24 colors.");
  const cleaned = colors.map((color) => {
    const value = String(color).trim();
    if (!/^#[0-9a-f]{6}$/i.test(value)) throw new Error(`“${value}” is not a #RRGGBB color.`);
    return value.toUpperCase();
  });
  const familyPalettes = { ...getDeckSettings().familyPalettes, [family as PaletteFamily]: cleaned };
  updateDeckSettings({ familyPalettes });
  return { ok: true, message: `${family} palette saved (${cleaned.length} colors).`, charts: countCharts() };
}

export function resetFamilyPalette(family: string): { ok: true; message: string } {
  if (!PALETTE_FAMILIES.includes(family as PaletteFamily)) throw new Error(`Unknown chart family: ${family}`);
  const familyPalettes = { ...getDeckSettings().familyPalettes };
  delete familyPalettes[family as PaletteFamily];
  updateDeckSettings({ familyPalettes });
  return { ok: true, message: `${family} palette reset to the current color theme.` };
}
