import {
  CHART_META_PREFIX, decodeMetadata, type ChartMetadata,
} from "../core/chart-data";

const DECK_KEY = "slideAid.deck.v1";
const CHART_KEY_PREFIX = "slideAid.chart.v2.";
const CHART_REF_PREFIX = "slide-aid-chart:";
// Apps Script properties are limited by UTF-8 byte size. Two thousand UTF-16
// code units remain below the 9 KB/property limit even for four-byte Unicode.
const CHUNK_SIZE = 2000;

export interface LayoutSlot {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface LayoutPreset {
  name: string;
  slots: LayoutSlot[];
  createdAt: string;
}

export interface DeckSettings {
  palette?: string;
  libraryPresentationId?: string;
  libraryPresentationUrl?: string;
  layouts: LayoutPreset[];
}

const DEFAULT_DECK_SETTINGS: DeckSettings = { layouts: [] };

function documentProperties(): GoogleAppsScript.Properties.Properties {
  const properties = PropertiesService.getDocumentProperties();
  if (!properties) throw new Error("Shared document storage is unavailable. Run Slide Aid as an installed editor add-on.");
  return properties;
}

function withDocumentLock<T>(operation: () => T): T {
  const lock = LockService.getDocumentLock();
  if (!lock) return operation();
  lock.waitLock(5000);
  try { return operation(); } finally { lock.releaseLock(); }
}

function writeChunkedUnlocked(key: string, value: string): void {
  const properties = documentProperties();
  const oldCount = Number(properties.getProperty(`${key}.chunks`) ?? 0);
  const chunks: Record<string, string> = {};
  const count = Math.max(1, Math.ceil(value.length / CHUNK_SIZE));
  for (let index = 0; index < count; index += 1) chunks[`${key}.${index}`] = value.slice(index * CHUNK_SIZE, (index + 1) * CHUNK_SIZE);
  chunks[`${key}.chunks`] = String(count);
  properties.setProperties(chunks, false);
  for (let index = count; index < oldCount; index += 1) properties.deleteProperty(`${key}.${index}`);
}

function writeChunked(key: string, value: string): void {
  withDocumentLock(() => writeChunkedUnlocked(key, value));
}

function readChunked(key: string): string | null {
  const properties = documentProperties();
  const count = Number(properties.getProperty(`${key}.chunks`) ?? 0);
  if (!count) return null;
  let value = "";
  for (let index = 0; index < count; index += 1) {
    const chunk = properties.getProperty(`${key}.${index}`);
    if (chunk == null) return null;
    value += chunk;
  }
  return value;
}

function deleteChunked(key: string): void {
  withDocumentLock(() => {
    const properties = documentProperties();
    const count = Number(properties.getProperty(`${key}.chunks`) ?? 0);
    for (let index = 0; index < count; index += 1) properties.deleteProperty(`${key}.${index}`);
    properties.deleteProperty(`${key}.chunks`);
  });
}

export function getDeckSettings(): DeckSettings {
  const raw = readChunked(DECK_KEY);
  if (!raw) return { ...DEFAULT_DECK_SETTINGS, layouts: [] };
  try {
    const parsed = JSON.parse(raw) as Partial<DeckSettings>;
    return { ...DEFAULT_DECK_SETTINGS, ...parsed, layouts: Array.isArray(parsed.layouts) ? parsed.layouts : [] };
  } catch {
    return { ...DEFAULT_DECK_SETTINGS, layouts: [] };
  }
}

export function updateDeckSettings(patch: Partial<DeckSettings>): DeckSettings {
  return withDocumentLock(() => {
    const next = { ...getDeckSettings(), ...patch };
    writeChunkedUnlocked(DECK_KEY, JSON.stringify(next));
    return next;
  });
}

export function chartDescription(metadata: ChartMetadata): string {
  return `Chart Aid ${metadata.kind} chart. Editable with Slide Aid. [${CHART_REF_PREFIX}${metadata.id}]`;
}

export function saveChartMetadata(metadata: ChartMetadata): void {
  writeChunked(`${CHART_KEY_PREFIX}${metadata.id}`, JSON.stringify(metadata));
}

export function deleteChartMetadata(id: string): void {
  deleteChunked(`${CHART_KEY_PREFIX}${id}`);
}

export function chartIdFromDescription(description: string): string | null {
  const match = description.match(/\[slide-aid-chart:([A-Za-z0-9_-]+)\]/);
  return match?.[1] ?? null;
}

export function loadChartMetadata(description: string): ChartMetadata | null {
  const id = chartIdFromDescription(description);
  if (id) {
    const raw = readChunked(`${CHART_KEY_PREFIX}${id}`);
    if (!raw) return null;
    try {
      return JSON.parse(raw) as ChartMetadata;
    } catch {
      return null;
    }
  }

  // v0.1 stored the complete payload in alt text. Read it once, migrate it to
  // shared document properties, and let the next rebuild replace the alt text.
  if (description.startsWith(CHART_META_PREFIX)) {
    const legacy = decodeMetadata(description);
    if (legacy) saveChartMetadata(legacy);
    return legacy;
  }
  return null;
}
