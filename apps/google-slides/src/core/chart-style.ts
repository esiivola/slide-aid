import type { ChartKind } from "./chart-data";

// Port of the PowerPoint Chart Aid style system (modChartStyle.bas). The key
// names, defaults, per-type override rule and control lists are deliberately
// identical, so a setting means the same thing in both products and the shared
// contract in shared/specs stays honest.
//
// PowerPoint persists this as chartstyle.txt lines ("COL.ClusterFill=0.8") in
// its sandbox folder. Here it is a flat map in Document Properties, so the whole
// deck's collaborators see the same chart styling.

export type StyleStore = Record<string, string>;

export interface StyleKeyDef {
  key: string;
  value: string;
  description: string;
}

/** Key, default, description - single source of truth, mirroring KeyDefs(). */
export const STYLE_KEYS: readonly StyleKeyDef[] = [
  { key: "PlotWidthCm", value: "12", description: "Default chart width (cm)" },
  { key: "PlotHeightCm", value: "8", description: "Default chart height (cm)" },
  { key: "ClusterFill", value: "0.72", description: "Bar group width as share of category slot (clustered)" },
  { key: "StackFill", value: "0.65", description: "Bar width as share of slot (stacked / 100%)" },
  { key: "WaterfallFill", value: "0.62", description: "Bar width as share of slot (waterfall)" },
  { key: "MekkoGapPt", value: "2", description: "Gap between Mekko columns (pt)" },
  { key: "LabelSizePt", value: "9", description: "Font size of chart labels (pt)" },
  { key: "Decimals", value: "auto", description: "Value decimals: auto, 0, 1 or 2" },
  { key: "ValueLabels", value: "1", description: "Show value labels (1 = yes, 0 = no)" },
  { key: "TotalLabels", value: "1", description: "Show totals on stacked columns (1/0)" },
  { key: "Legend", value: "1", description: "Show the series legend (1/0)" },
  { key: "MarkerSizePt", value: "5", description: "Line-chart marker size (pt)" },
  { key: "WaterfallUp", value: "9BBB59", description: "Waterfall: positive segments (hex)" },
  { key: "WaterfallDown", value: "C0504D", description: "Waterfall: negative segments (hex)" },
  { key: "WaterfallTotal", value: "BFBFBF", description: "Waterfall: subtotal bars (hex)" },
  { key: "GanttBarColor", value: "theme", description: "Gantt bars: hex color, or 'theme'" },
];

const DEFAULTS: Record<string, string> = STYLE_KEYS.reduce<Record<string, string>>((all, def) => {
  all[def.key.toLowerCase()] = def.value;
  return all;
}, {});

export function isKnownStyleKey(key: string): boolean {
  const base = key.includes(".") ? key.slice(key.indexOf(".") + 1) : key;
  return base.toLowerCase() in DEFAULTS;
}

function rawLookup(store: StyleStore, key: string): string | undefined {
  const wanted = key.toLowerCase();
  for (const name of Object.keys(store)) {
    if (name.toLowerCase() === wanted) return store[name];
  }
  return undefined;
}

/** A per-type override ("COL.ClusterFill") wins over the global key. */
export function styleString(store: StyleStore, kind: string | null, key: string): string {
  if (kind) {
    const scoped = rawLookup(store, `${kind}.${key}`);
    if (scoped !== undefined && scoped !== "") return scoped;
  }
  const global = rawLookup(store, key);
  if (global !== undefined && global !== "") return global;
  return DEFAULTS[key.toLowerCase()] ?? "";
}

export function styleNumber(store: StyleStore, kind: string | null, key: string): number {
  const parsed = Number.parseFloat(styleString(store, kind, key).replace(",", "."));
  return Number.isFinite(parsed) ? parsed : Number.parseFloat(DEFAULTS[key.toLowerCase()] ?? "0") || 0;
}

export function styleFlag(store: StyleStore, kind: string | null, key: string): boolean {
  const value = styleString(store, kind, key).trim().toLowerCase();
  return !(value === "0" || value === "false" || value === "no" || value === "");
}

/** Normalizes a stored color to #RRGGBB, or null when it means "follow the theme". */
export function styleColor(store: StyleStore, kind: string | null, key: string): string | null {
  const value = styleString(store, kind, key).trim();
  if (!value || value.toLowerCase() === "theme") return null;
  const hex = value.startsWith("#") ? value.slice(1) : value;
  return /^[0-9a-f]{6}$/i.test(hex) ? `#${hex.toUpperCase()}` : null;
}

function withThousands(value: string): string {
  const [whole, fraction] = value.split(".");
  const sign = whole!.startsWith("-") ? "-" : "";
  const digits = sign ? whole!.slice(1) : whole!;
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return `${sign}${grouped}${fraction ? `.${fraction}` : ""}`;
}

/**
 * Mirrors FmtNum(): explicit 0/1/2 decimals, or "auto" - whole numbers print
 * without decimals and everything else gets one.
 */
export function formatValue(value: number, decimals: string): string {
  const mode = decimals.trim().toLowerCase();
  if (mode === "0" || mode === "1" || mode === "2") return withThousands(value.toFixed(Number(mode)));
  const whole = Math.abs(value - Math.round(value)) < 0.000001;
  return withThousands(value.toFixed(whole ? 0 : 1));
}

// ---------------------------------------------------------------------------
// Per-family palettes. PowerPoint keeps three palette files and picks one by
// chart family, so recoloring "Lines" never disturbs the bar charts.
// ---------------------------------------------------------------------------

export type PaletteFamily = "BARS" | "LINES" | "PIES";

export const PALETTE_FAMILIES: readonly PaletteFamily[] = ["BARS", "LINES", "PIES"];

export const FAMILY_LABELS: Record<PaletteFamily, string> = {
  BARS: "Bars (column, bar, stacked, 100%, waterfall, Mekko, Gantt)",
  LINES: "Lines (line, area, scatter, bubble)",
  PIES: "Pies (pie, doughnut)",
};

export function familyForKind(kind: ChartKind | string): PaletteFamily {
  if (["LINE", "AREA", "SCAT", "BUB"].includes(kind)) return "LINES";
  if (["PIE", "DON"].includes(kind)) return "PIES";
  return "BARS";
}

// ---------------------------------------------------------------------------
// Control definitions for the Chart Settings panel. These are the same lists
// PowerPoint feeds to its native NSAlert panel, curated to the parameters each
// builder actually reads, so the two panels offer the same choices in the same
// order.
// ---------------------------------------------------------------------------

export type ControlType = "num" | "pct" | "check" | "popup" | "color";

export interface StyleControl {
  type: ControlType;
  key: string;
  label: string;
  min?: number;
  max?: number;
  options?: string[];
}

const GLOBAL_CONTROLS: readonly StyleControl[] = [
  { type: "num", key: "PlotWidthCm", label: "Default width (cm)", min: 4, max: 30 },
  { type: "num", key: "PlotHeightCm", label: "Default height (cm)", min: 3, max: 20 },
  { type: "num", key: "LabelSizePt", label: "Label size (pt)", min: 5, max: 24 },
  { type: "popup", key: "Decimals", label: "Decimals", options: ["auto", "0", "1", "2"] },
  { type: "check", key: "ValueLabels", label: "Show value labels" },
  { type: "check", key: "TotalLabels", label: "Show totals on stacked columns" },
  { type: "check", key: "Legend", label: "Show legend (charts with 2+ series)" },
];

const LABEL_SIZE: StyleControl = { type: "num", key: "LabelSizePt", label: "Label size (pt)", min: 5, max: 24 };
const DECIMALS: StyleControl = { type: "popup", key: "Decimals", label: "Decimals", options: ["auto", "0", "1", "2"] };

export function controlsFor(scope: string): readonly StyleControl[] {
  switch (scope) {
    case "GLOBAL":
      return GLOBAL_CONTROLS;
    case "COL":
    case "BAR":
      return [
        { type: "pct", key: "ClusterFill", label: "Bar width (%)", min: 30, max: 100 },
        { type: "check", key: "ValueLabels", label: "Show value labels" },
        { type: "check", key: "Legend", label: "Show legend (2+ series)" },
        LABEL_SIZE, DECIMALS,
      ];
    case "STK":
    case "SBR":
    case "PCT":
      return [
        { type: "pct", key: "StackFill", label: "Bar width (%)", min: 30, max: 100 },
        { type: "check", key: "ValueLabels", label: "Show segment labels" },
        { type: "check", key: "TotalLabels", label: "Show totals" },
        { type: "check", key: "Legend", label: "Show legend (2+ series)" },
        LABEL_SIZE, DECIMALS,
      ];
    case "WF":
      return [
        { type: "pct", key: "WaterfallFill", label: "Bar width (%)", min: 30, max: 100 },
        { type: "color", key: "WaterfallUp", label: "Positive color" },
        { type: "color", key: "WaterfallDown", label: "Negative color" },
        { type: "color", key: "WaterfallTotal", label: "Subtotal color" },
        LABEL_SIZE, DECIMALS,
      ];
    case "MEK":
      return [
        { type: "num", key: "MekkoGapPt", label: "Column gap (pt)", min: 0, max: 20 },
        LABEL_SIZE, DECIMALS,
      ];
    case "LINE":
      return [
        { type: "num", key: "MarkerSizePt", label: "Marker size (pt)", min: 0, max: 12 },
        { type: "check", key: "ValueLabels", label: "Show value labels" },
        { type: "check", key: "Legend", label: "Show legend (2+ series)" },
        LABEL_SIZE, DECIMALS,
      ];
    case "AREA":
      return [LABEL_SIZE];
    case "GANTT":
      return [
        { type: "color", key: "GanttBarColor", label: "Bar color (blank follows theme)" },
        LABEL_SIZE,
      ];
    default:
      // PIE, DON, SCAT, BUB
      return [LABEL_SIZE, DECIMALS];
  }
}

export function scopeLabel(scope: string): string {
  const names: Record<string, string> = {
    GLOBAL: "New-chart defaults", COL: "Column", BAR: "Bar", STK: "Stacked", SBR: "Stacked bar",
    PCT: "100%", WF: "Waterfall", MEK: "Mekko", LINE: "Line", AREA: "Area", PIE: "Pie",
    DON: "Doughnut", SCAT: "Scatter", BUB: "Bubble", GANTT: "Gantt",
  };
  return names[scope] ?? scope;
}

/** Values the panel should show: the effective setting for every control in scope. */
export function controlValues(store: StyleStore, scope: string): Record<string, string> {
  const kind = scope === "GLOBAL" ? null : scope;
  const values: Record<string, string> = {};
  for (const control of controlsFor(scope)) {
    if (control.type === "pct") values[control.key] = String(Math.round(styleNumber(store, kind, control.key) * 100));
    else if (control.type === "check") values[control.key] = styleFlag(store, kind, control.key) ? "1" : "0";
    else if (control.type === "color") values[control.key] = styleColor(store, kind, control.key) ?? "";
    else values[control.key] = styleString(store, kind, control.key);
  }
  return values;
}

/**
 * Validates a panel submission and turns it into store entries. Scoped writes
 * are prefixed with the chart kind, which is what gives PowerPoint's
 * "this chart type only" behavior.
 */
export function applyControlValues(scope: string, submitted: Record<string, unknown>): StyleStore {
  const patch: StyleStore = {};
  for (const control of controlsFor(scope)) {
    if (!(control.key in submitted)) continue;
    const raw = String(submitted[control.key] ?? "").trim();
    const name = scope === "GLOBAL" ? control.key : `${scope}.${control.key}`;
    if (control.type === "check") {
      patch[name] = raw === "1" || raw.toLowerCase() === "true" ? "1" : "0";
      continue;
    }
    if (control.type === "popup") {
      if (!control.options?.includes(raw)) throw new Error(`${control.label} must be one of ${control.options?.join(", ")}.`);
      patch[name] = raw;
      continue;
    }
    if (control.type === "color") {
      if (!raw || raw.toLowerCase() === "theme") {
        patch[name] = "theme";
        continue;
      }
      const hex = raw.startsWith("#") ? raw.slice(1) : raw;
      if (!/^[0-9a-f]{6}$/i.test(hex)) throw new Error(`${control.label} must be a #RRGGBB color.`);
      patch[name] = hex.toUpperCase();
      continue;
    }
    const value = Number.parseFloat(raw.replace(",", "."));
    if (!Number.isFinite(value)) throw new Error(`${control.label} must be a number.`);
    const min = control.min ?? Number.NEGATIVE_INFINITY;
    const max = control.max ?? Number.POSITIVE_INFINITY;
    if (value < min || value > max) throw new Error(`${control.label} must be between ${min} and ${max}.`);
    patch[name] = control.type === "pct" ? String(Math.round(value) / 100) : String(value);
  }
  return patch;
}

/** Drops every override for one scope so it inherits the global defaults again. */
export function clearScope(store: StyleStore, scope: string): StyleStore {
  const next: StyleStore = {};
  for (const [key, value] of Object.entries(store)) {
    if (scope === "GLOBAL") {
      if (key.includes(".")) next[key] = value;
      continue;
    }
    if (!key.toLowerCase().startsWith(`${scope.toLowerCase()}.`)) next[key] = value;
  }
  return next;
}
