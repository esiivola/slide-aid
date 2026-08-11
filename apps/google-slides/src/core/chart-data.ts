export type ChartKind =
  | "COL" | "BAR" | "STK" | "SBR" | "PCT" | "MEK" | "WF"
  | "LINE" | "AREA" | "PIE" | "DON" | "SCAT" | "BUB" | "GANTT";

export interface ChartData {
  cells: string[][];
}

export interface ChartRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Where the value axis sits after a build, so annotations can be placed on the
 * chart's own scale instead of guessing from pixels. This is the Slides
 * equivalent of PowerPoint's SACH_Y0 / SACH_PPU shape tags.
 */
export interface ChartAxis {
  /** Coordinate of value zero: an x for horizontal charts, a y otherwise. */
  zero: number;
  /** Points per data unit. */
  scale: number;
  horizontal: boolean;
  /** Start and length of the plot band across the category direction. */
  plotStart: number;
  plotSize: number;
}

export interface ChartMetadata {
  schema: 1;
  id: string;
  kind: ChartKind;
  data: ChartData;
  rect: ChartRect;
  palette: string;
  overrides: Record<string, string>;
  source?: SheetSource;
  axis?: ChartAxis;
}

const BAR_TAG = /\[slide-aid-bar:(\d+):(\d+):(-?[\d.]+)\]/;

export interface BarReference {
  series: number;
  category: number;
  value: number;
}

/** Marks one drawn bar/segment/point with the data it came from. */
export function barTag(series: number, category: number, value: number): string {
  return `[slide-aid-bar:${series}:${category}:${value}]`;
}

export function parseBarTag(description: string): BarReference | null {
  const match = description.match(BAR_TAG);
  if (!match) return null;
  const value = Number.parseFloat(match[3]!);
  if (!Number.isFinite(value)) return null;
  return { series: Number(match[1]), category: Number(match[2]), value };
}

/** Compound annual growth rate between two values over a number of periods. */
export function cagr(first: number, last: number, periods: number): number {
  if (!(periods > 0)) throw new Error("CAGR needs at least one period.");
  if (first <= 0 || last <= 0) throw new Error("CAGR needs two positive values.");
  return ((last / first) ** (1 / periods) - 1) * 100;
}

export function percentDifference(first: number, last: number): number {
  if (first === 0) throw new Error("Percent difference needs a non-zero starting value.");
  return ((last - first) / Math.abs(first)) * 100;
}

export interface SheetSource {
  spreadsheetId: string;
  spreadsheetUrl: string;
  sheetName: string;
  rangeA1: string;
}

export const CHART_META_PREFIX = "SLIDE_AID_CHART_V1:";
export const DATASHEET_PREFIX = "SLIDE_AID_DATASHEET_V1:";

export function parseNumber(value: string): number {
  const normalized = value.trim().replace(/\s/g, "").replace(",", ".").replace("%", "");
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function isSubtotal(value: string): boolean {
  return ["e", "="].includes(value.trim().toLowerCase());
}

export function validateChartData(kind: ChartKind, data: ChartData): void {
  if (data.cells.length === 0 || data.cells[0]?.length === 0) throw new Error("The selected table is empty.");
  const columns = Math.max(...data.cells.map((row) => row.length));
  if (["COL", "BAR", "STK", "SBR", "PCT", "MEK", "LINE", "AREA"].includes(kind)) {
    if (data.cells.length < 2 || columns < 2) throw new Error("This chart needs a header row, series names and values.");
  }
  if (["PIE", "DON", "WF"].includes(kind) && columns < 2) throw new Error("This chart needs rows of label and value.");
  if (["SCAT", "BUB"].includes(kind) && columns < 3) throw new Error("Scatter and bubble charts need label, x and y columns.");
  if (kind === "GANTT" && columns < 3) throw new Error("Gantt needs activity, start and end columns.");
}

export function encodeMetadata(metadata: ChartMetadata): string {
  return CHART_META_PREFIX + JSON.stringify(metadata);
}

export function decodeMetadata(description: string): ChartMetadata | null {
  if (!description.startsWith(CHART_META_PREFIX)) return null;
  try {
    const value = JSON.parse(description.slice(CHART_META_PREFIX.length)) as ChartMetadata;
    if (value.schema !== 1 || !value.id || !value.kind || !Array.isArray(value.data?.cells)) return null;
    return value;
  } catch {
    return null;
  }
}

export function paletteColor(palette: readonly string[], index: number, overrides: Record<string, string> = {}): string {
  const override = overrides[String(index + 1)];
  return override ?? palette[index % palette.length] ?? "#4472C4";
}
