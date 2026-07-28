import {
  DATASHEET_PREFIX, isSubtotal, paletteColor,
  parseNumber, validateChartData, type ChartData, type ChartKind, type ChartMetadata, type ChartRect, type SheetSource,
} from "../core/chart-data";
import { activeContext, elementBox } from "../slides/selection";
import { currentPalette, getSettings, PALETTES, updateSettings } from "../storage/preferences";
import { chartDescription, getDeckSettings, loadChartMetadata, saveChartMetadata, updateDeckSettings } from "../storage/document-state";
import { readSheetData, refreshSheetData } from "../integrations/sheets";
import { ChartBatch } from "../slides/chart-batch";

type PageElement = GoogleAppsScript.Slides.PageElement;
type Slide = GoogleAppsScript.Slides.Slide;

const GREY = "#595959";
const LIGHT_GREY = "#BFBFBF";
const GREEN = "#9BBB59";
const RED = "#C0504D";

interface BuildContext {
  slide: Slide;
  rect: ChartRect;
  batch: ChartBatch;
  palette: readonly string[];
  metadata: ChartMetadata;
}

function addRect(ctx: BuildContext, left: number, top: number, width: number, height: number, color: string): void {
  ctx.batch.addShape("RECTANGLE", left, top, Math.max(0.1, width), Math.max(0.1, height), color);
}

function addEllipse(ctx: BuildContext, left: number, top: number, width: number, height: number, color: string): void {
  ctx.batch.addShape("ELLIPSE", left, top, Math.max(0.1, width), Math.max(0.1, height), color);
}

function addLine(ctx: BuildContext, x1: number, y1: number, x2: number, y2: number, color = GREY, weight = 1): void {
  ctx.batch.addLine(x1, y1, x2, y2, color, weight);
}

function addLabel(
  ctx: BuildContext, text: string, left: number, top: number, width: number, height = 14,
  size = 9, alignment: GoogleAppsScript.Slides.ParagraphAlignment = SlidesApp.ParagraphAlignment.CENTER,
  color = "#404040",
): void {
  const apiAlignment = alignment === SlidesApp.ParagraphAlignment.END ? "END" : alignment === SlidesApp.ParagraphAlignment.START ? "START" : "CENTER";
  ctx.batch.addText(text, left, top, Math.max(1, width), Math.max(1, height), size, apiAlignment, color);
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

function drawColumns(ctx: BuildContext, stacked: boolean, normalized: boolean, horizontal: boolean): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const categoryCount = categories.length;
  if (!categoryCount || !series.length) return;
  const values = series.flatMap((item) => item.values);
  const totals = categories.map((_, c) => series.reduce((sum, item) => sum + Math.abs(item.values[c] ?? 0), 0));
  const range = normalized ? { min: 0, max: 100 } : numericRange(stacked ? totals : values);
  const labelSpace = horizontal ? 70 : 22;
  const plotLeft = ctx.rect.left + (horizontal ? labelSpace : 4);
  const plotTop = ctx.rect.top + 8;
  const plotWidth = ctx.rect.width - (horizontal ? labelSpace + 8 : 8);
  const plotHeight = ctx.rect.height - (horizontal ? 12 : labelSpace + 8);
  const scale = (horizontal ? plotWidth : plotHeight) / (range.max - range.min);
  const baseline = horizontal ? plotLeft - range.min * scale : plotTop + range.max * scale;
  addLine(ctx, horizontal ? baseline : plotLeft, horizontal ? plotTop : baseline, horizontal ? baseline : plotLeft + plotWidth, horizontal ? plotTop + plotHeight : baseline);

  const slot = (horizontal ? plotHeight : plotWidth) / categoryCount;
  const clusterFill = stacked ? 0.65 : 0.72;
  categories.forEach((category, categoryIndex) => {
    let positive = 0;
    let negative = 0;
    series.forEach((item, seriesIndex) => {
      const raw = item.values[categoryIndex] ?? 0;
      const value = normalized ? (totals[categoryIndex] ? (raw / totals[categoryIndex]!) * 100 : 0) : raw;
      const color = paletteColor(ctx.palette, seriesIndex, ctx.metadata.overrides);
      const thickness = stacked ? slot * clusterFill : (slot * clusterFill) / series.length;
      const offset = stacked ? (slot - thickness) / 2 : (slot - slot * clusterFill) / 2 + seriesIndex * thickness;
      const start = value >= 0 ? positive : negative;
      if (stacked) {
        if (value >= 0) positive += value;
        else negative += value;
      }
      if (horizontal) {
        const x = baseline + (stacked ? start : 0) * scale;
        const w = value * scale;
        addRect(ctx, w >= 0 ? x : x + w, plotTop + categoryIndex * slot + offset, Math.abs(w), thickness, color);
      } else {
        const y0 = baseline - (stacked ? start : 0) * scale;
        const h = value * scale;
        addRect(ctx, plotLeft + categoryIndex * slot + offset, h >= 0 ? y0 - h : y0, thickness, Math.abs(h), color);
      }
    });
    if (horizontal) addLabel(ctx, category, ctx.rect.left, plotTop + categoryIndex * slot + slot / 2 - 7, labelSpace - 4, 14, 9, SlidesApp.ParagraphAlignment.END);
    else addLabel(ctx, category, plotLeft + categoryIndex * slot, plotTop + plotHeight + 2, slot, 14);
  });
}

function drawWaterfall(ctx: BuildContext): void {
  const rows = ctx.metadata.data.cells;
  const values: { label: string; value: number; total: boolean; base: number }[] = [];
  let running = 0;
  for (const row of rows) {
    const total = isSubtotal(row[1] ?? "");
    const value = total ? running : parseNumber(row[1] ?? "");
    values.push({ label: row[0] ?? "", value, total, base: total ? 0 : running });
    if (!total) running += value;
  }
  const extrema = values.flatMap((item) => [item.base, item.total ? item.value : item.base + item.value]);
  const range = numericRange(extrema);
  const plotTop = ctx.rect.top + 8;
  const plotHeight = ctx.rect.height - 30;
  const scale = plotHeight / (range.max - range.min);
  const baseline = plotTop + range.max * scale;
  const slot = ctx.rect.width / Math.max(values.length, 1);
  addLine(ctx, ctx.rect.left, baseline, ctx.rect.left + ctx.rect.width, baseline);
  values.forEach((item, index) => {
    const end = item.total ? item.value : item.base + item.value;
    const start = item.total ? 0 : item.base;
    const y1 = baseline - start * scale;
    const y2 = baseline - end * scale;
    const color = item.total ? LIGHT_GREY : item.value >= 0 ? GREEN : RED;
    addRect(ctx, ctx.rect.left + index * slot + slot * 0.19, Math.min(y1, y2), slot * 0.62, Math.max(0.5, Math.abs(y2 - y1)), color);
    addLabel(ctx, item.label, ctx.rect.left + index * slot, plotTop + plotHeight + 2, slot, 14, 8);
    if (!item.total && index < values.length - 1) addLine(ctx, ctx.rect.left + (index + 0.81) * slot, y2, ctx.rect.left + (index + 1.19) * slot, y2, "#A0A0A0", 0.75);
  });
}

function drawMekko(ctx: BuildContext): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const totals = categories.map((_, c) => series.reduce((sum, item) => sum + Math.max(0, item.values[c] ?? 0), 0));
  const all = totals.reduce((sum, value) => sum + value, 0) || 1;
  let x = ctx.rect.left;
  categories.forEach((category, c) => {
    const width = (totals[c]! / all) * ctx.rect.width;
    let y = ctx.rect.top + ctx.rect.height - 18;
    series.forEach((item, s) => {
      const height = totals[c] ? ((item.values[c] ?? 0) / totals[c]!) * (ctx.rect.height - 24) : 0;
      y -= height;
      addRect(ctx, x + 1, y, Math.max(0.5, width - 2), Math.max(0.5, height), paletteColor(ctx.palette, s, ctx.metadata.overrides));
    });
    addLabel(ctx, category, x, ctx.rect.top + ctx.rect.height - 14, width, 14, 8);
    x += width;
  });
}

function drawLineChart(ctx: BuildContext): void {
  const { categories, series } = gridData(ctx.metadata.data);
  const values = series.flatMap((item) => item.values);
  const range = numericRange(values);
  const plotLeft = ctx.rect.left + 10;
  const plotTop = ctx.rect.top + 8;
  const plotWidth = ctx.rect.width - 20;
  const plotHeight = ctx.rect.height - 30;
  const scale = plotHeight / (range.max - range.min);
  const baseline = plotTop + range.max * scale;
  addLine(ctx, plotLeft, baseline, plotLeft + plotWidth, baseline);
  const slot = plotWidth / Math.max(categories.length, 1);
  series.forEach((item, s) => {
    const color = paletteColor(ctx.palette, s, ctx.metadata.overrides);
    item.values.forEach((value, c) => {
      const x = plotLeft + c * slot + slot / 2;
      const y = baseline - value * scale;
      if (c > 0) {
        const previous = item.values[c - 1] ?? 0;
        addLine(ctx, x - slot, baseline - previous * scale, x, y, color, 2);
      }
      addEllipse(ctx, x - 2.5, y - 2.5, 5, 5, color);
    });
  });
  categories.forEach((category, c) => addLabel(ctx, category, plotLeft + c * slot, plotTop + plotHeight + 2, slot, 14, 8));
}

function drawScatter(ctx: BuildContext, bubble: boolean): void {
  const rows = ctx.metadata.data.cells;
  const points = rows.map((row) => ({ label: row[0] ?? "", x: parseNumber(row[1] ?? ""), y: parseNumber(row[2] ?? ""), size: parseNumber(row[3] ?? "") }));
  const xr = numericRange(points.map((point) => point.x));
  const yr = numericRange(points.map((point) => point.y));
  const maxSize = Math.max(1, ...points.map((point) => point.size));
  const left = ctx.rect.left + 30;
  const top = ctx.rect.top + 8;
  const width = ctx.rect.width - 40;
  const height = ctx.rect.height - 28;
  addLine(ctx, left, top, left, top + height);
  addLine(ctx, left, top + height, left + width, top + height);
  points.forEach((point, index) => {
    const x = left + ((point.x - xr.min) / (xr.max - xr.min)) * width;
    const y = top + height - ((point.y - yr.min) / (yr.max - yr.min)) * height;
    const diameter = bubble ? 6 + 22 * Math.sqrt(Math.max(0, point.size) / maxSize) : 8;
    addEllipse(ctx, x - diameter / 2, y - diameter / 2, diameter, diameter, paletteColor(ctx.palette, index, ctx.metadata.overrides));
    addLabel(ctx, point.label, x + diameter / 2 + 2, y - 7, 70, 14, 8, SlidesApp.ParagraphAlignment.START);
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
  const rows = ctx.metadata.data.cells.map((row) => ({ label: row[0] ?? "", start: dateOrNumber(row[1] ?? ""), end: dateOrNumber(row[2] ?? "") }));
  const min = Math.min(...rows.map((row) => row.start));
  const max = Math.max(...rows.map((row) => row.end));
  const span = Math.max(1, max - min);
  const labelWidth = Math.min(110, ctx.rect.width * 0.32);
  const left = ctx.rect.left + labelWidth;
  const width = ctx.rect.width - labelWidth - 4;
  const rowHeight = (ctx.rect.height - 8) / Math.max(rows.length, 1);
  rows.forEach((row, index) => {
    addLabel(ctx, row.label, ctx.rect.left, ctx.rect.top + index * rowHeight + rowHeight / 2 - 7, labelWidth - 4, 14, 9, SlidesApp.ParagraphAlignment.END);
    const x = left + ((row.start - min) / span) * width;
    const end = left + ((row.end - min) / span) * width;
    if (Math.abs(row.end - row.start) < 1e-9) {
      ctx.batch.addShape("DIAMOND", x - 5, ctx.rect.top + index * rowHeight + rowHeight / 2 - 5, 10, 10, paletteColor(ctx.palette, 0, ctx.metadata.overrides));
    } else addRect(ctx, x, ctx.rect.top + index * rowHeight + rowHeight * 0.25, Math.max(1, end - x), rowHeight * 0.5, paletteColor(ctx.palette, 0, ctx.metadata.overrides));
  });
}

function staticChart(ctx: BuildContext, kind: "AREA" | "PIE" | "DON"): PageElement {
  const palette = ctx.palette.map((color, index) => paletteColor(ctx.palette, index, ctx.metadata.overrides));
  const pixelsWide = Math.max(240, Math.round(ctx.rect.width * 1.5));
  const pixelsHigh = Math.max(160, Math.round(ctx.rect.height * 1.5));
  if (kind === "AREA") {
    const { categories, series } = gridData(ctx.metadata.data);
    const builder = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Category");
    series.forEach((item) => builder.addColumn(Charts.ColumnType.NUMBER, item.name));
    categories.forEach((category, c) => builder.addRow([category, ...series.map((item) => item.values[c] ?? 0)]));
    const chart = Charts.newAreaChart().setDataTable(builder).setDimensions(pixelsWide, pixelsHigh)
      .setColors([...palette]).setStacked().setBackgroundColor("white").setLegendPosition(Charts.Position.RIGHT).build();
    return ctx.slide.insertImage(chart.getBlob(), ctx.rect.left, ctx.rect.top, ctx.rect.width, ctx.rect.height) as unknown as PageElement;
  }
  const builder = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Label").addColumn(Charts.ColumnType.NUMBER, "Value");
  ctx.metadata.data.cells.forEach((row) => builder.addRow([row[0] ?? "", Math.abs(parseNumber(row[1] ?? ""))]));
  const pie = Charts.newPieChart().setDataTable(builder).setDimensions(pixelsWide, pixelsHigh).setColors([...palette])
    .setBackgroundColor("white").setLegendPosition(Charts.Position.RIGHT);
  if (kind === "DON") pie.setOption("pieHole", 0.45);
  return ctx.slide.insertImage(pie.build().getBlob(), ctx.rect.left, ctx.rect.top, ctx.rect.width, ctx.rect.height) as unknown as PageElement;
}

function tagChart(element: PageElement, metadata: ChartMetadata): void {
  saveChartMetadata(metadata);
  element.setTitle(`Slide Aid ${metadata.kind} chart`);
  element.setDescription(chartDescription(metadata));
  element.select();
}

function drawChart(slide: Slide, metadata: ChartMetadata): void {
  validateChartData(metadata.kind, metadata.data);
  const palette = PALETTES[metadata.palette] ?? currentPalette();
  if (["AREA", "PIE", "DON"].includes(metadata.kind)) {
    const image = staticChart({ slide, rect: metadata.rect, batch: new ChartBatch(SlidesApp.getActivePresentation().getId(), slide.getObjectId()), palette, metadata }, metadata.kind as "AREA" | "PIE" | "DON");
    tagChart(image, metadata);
    return;
  }
  const presentation = SlidesApp.getActivePresentation();
  const ctx: BuildContext = { slide, rect: metadata.rect, batch: new ChartBatch(presentation.getId(), slide.getObjectId()), palette, metadata };
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
  saveChartMetadata(metadata);
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
  const width = 12 * 28.3464567;
  const height = 8 * 28.3464567;
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

export function restyleCharts(allCharts: boolean): { ok: true; message: string } {
  if (!allCharts) return rebuildChart();
  const context = activeContext();
  const records: { slide: Slide; element: PageElement; metadata: ChartMetadata }[] = [];
  context.presentation.getSlides().forEach((slide) => slide.getPageElements().forEach((element) => {
    const metadata = loadChartMetadata(element.getDescription());
    if (metadata) records.push({ slide, element, metadata });
  }));
  records.forEach(({ slide, element, metadata }) => {
    const box = elementBox(element);
    const next = { ...metadata, palette: getDeckSettings().palette ?? getSettings().palette, rect: { left: box.left, top: box.top, width: box.width, height: box.height } };
    drawChart(slide, next);
    element.remove();
  });
  return { ok: true, message: `Restyled ${records.length} chart${records.length === 1 ? "" : "s"}.` };
}

export function setPalette(name: string): { ok: true; message: string } {
  if (!PALETTES[name]) throw new Error(`Unknown palette: ${name}`);
  updateSettings({ palette: name });
  updateDeckSettings({ palette: name });
  return { ok: true, message: `Palette set to ${name}. Rebuild or restyle charts to apply it.` };
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
