import type { Box } from "./geometry";
import type { LayoutSlot } from "../storage/document-state";

export function extractGoogleFileId(value: string): string {
  const trimmed = value.trim();
  const match = trimmed.match(/\/d\/([A-Za-z0-9_-]+)/) ?? trimmed.match(/^([A-Za-z0-9_-]{20,})$/);
  if (!match?.[1]) throw new Error("Enter a valid Google Slides or Sheets URL or file ID.");
  return match[1];
}

export interface LibraryReference {
  presentationId: string;
  slideId: string;
}

export function encodeLibraryReference(reference: LibraryReference): string {
  return `[slide-aid-library:${reference.presentationId}:${reference.slideId}]`;
}

export function decodeLibraryReference(description: string): LibraryReference | null {
  const match = description.match(/\[slide-aid-library:([A-Za-z0-9_-]+):([A-Za-z0-9_-]+)\]/);
  return match ? { presentationId: match[1]!, slideId: match[2]! } : null;
}

export function normalizeLayout(boxes: readonly Box[], slideWidth: number, slideHeight: number): LayoutSlot[] {
  if (!(slideWidth > 0) || !(slideHeight > 0)) throw new Error("Slide dimensions must be positive.");
  return boxes.map((box) => ({
    left: box.left / slideWidth,
    top: box.top / slideHeight,
    width: box.width / slideWidth,
    height: box.height / slideHeight,
  }));
}

export function projectLayout(slots: readonly LayoutSlot[], boxes: readonly Box[], slideWidth: number, slideHeight: number): Box[] {
  if (slots.length !== boxes.length) throw new Error(`This layout requires ${slots.length} selected objects.`);
  return boxes.map((box, index) => ({
    ...box,
    left: slots[index]!.left * slideWidth,
    top: slots[index]!.top * slideHeight,
    width: slots[index]!.width * slideWidth,
    height: slots[index]!.height * slideHeight,
  }));
}

export function isOutsideSlide(box: Box, slideWidth: number, slideHeight: number): boolean {
  return box.left < 0 || box.top < 0 || box.left + box.width > slideWidth || box.top + box.height > slideHeight;
}

function luminanceChannel(value: number): number {
  const channel = value / 255;
  return channel <= 0.03928 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
}

export function contrastRatio(first: string, second: string): number {
  const luminance = (hex: string): number => {
    if (!/^#[0-9a-f]{6}$/i.test(hex)) throw new Error("Contrast colors must use #RRGGBB format.");
    const channels = [1, 3, 5].map((start) => Number.parseInt(hex.slice(start, start + 2), 16));
    return 0.2126 * luminanceChannel(channels[0]!) + 0.7152 * luminanceChannel(channels[1]!) + 0.0722 * luminanceChannel(channels[2]!);
  };
  const a = luminance(first);
  const b = luminance(second);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

export interface IconLinePrimitive {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface IconShapePrimitive {
  kind: "rect" | "ellipse";
  x: number;
  y: number;
  width: number;
  height: number;
  filled: boolean;
}

export type IconPrimitive = IconLinePrimitive | IconShapePrimitive;

export interface IconPathElement {
  kind: "path";
  d: string;
  filled: boolean;
}

export interface IconPolylineElement {
  kind: "polyline";
  points: [number, number][];
  closed: boolean;
  filled: boolean;
}

export type IconElement = IconPrimitive | IconPathElement | IconPolylineElement;

export interface IconDefinition {
  id: string;
  name: string;
  category: string;
  aliases: string[];
  tags: string[];
  primitives: IconPrimitive[];
  elements: IconElement[];
}

function iconNumber(record: Record<string, unknown>, key: string): number {
  const value = record[key];
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 24) {
    throw new Error(`Icon primitive ${key} must be between 0 and 24.`);
  }
  return value;
}

function normalizePrimitive(primitive: unknown): IconPrimitive {
  if (!primitive || typeof primitive !== "object" || Array.isArray(primitive)) throw new Error("The icon contains an invalid primitive.");
  const item = primitive as Record<string, unknown>;
  if (item.kind === "line") {
    return {
      kind: "line",
      x1: iconNumber(item, "x1"),
      y1: iconNumber(item, "y1"),
      x2: iconNumber(item, "x2"),
      y2: iconNumber(item, "y2"),
    };
  }
  if (item.kind === "rect" || item.kind === "ellipse") {
    const width = iconNumber(item, "width");
    const height = iconNumber(item, "height");
    if (width <= 0 || height <= 0) throw new Error("Icon primitive dimensions must be positive.");
    return {
      kind: item.kind,
      x: iconNumber(item, "x"),
      y: iconNumber(item, "y"),
      width,
      height,
      filled: item.filled === true,
    };
  }
  throw new Error("The icon contains an unsupported primitive.");
}

function normalizeElement(element: unknown): IconElement {
  if (!element || typeof element !== "object" || Array.isArray(element)) throw new Error("The icon contains an invalid element.");
  const item = element as Record<string, unknown>;
  if (item.kind === "path") {
    if (typeof item.d !== "string" || !item.d.trim() || item.d.length > 400) throw new Error("The icon contains an invalid path.");
    const numbers = item.d.match(/-?\d+(?:\.\d+)?/g)?.map(Number) ?? [];
    if (!numbers.length || numbers.some((value) => !Number.isFinite(value) || value < 0 || value > 24)) {
      throw new Error("Icon path coordinates must be between 0 and 24.");
    }
    return { kind: "path", d: item.d.trim(), filled: item.filled === true };
  }
  if (item.kind === "polyline") {
    if (!Array.isArray(item.points) || item.points.length < 2) throw new Error("The icon contains an invalid polyline.");
    const points = item.points.map((point) => {
      if (!Array.isArray(point) || point.length !== 2) throw new Error("The icon contains an invalid polyline point.");
      const [x, y] = point;
      if (typeof x !== "number" || typeof y !== "number" || x < 0 || x > 24 || y < 0 || y > 24) {
        throw new Error("Icon polyline coordinates must be between 0 and 24.");
      }
      return [x, y] as [number, number];
    });
    return { kind: "polyline", points, closed: item.closed === true, filled: item.filled === true };
  }
  return normalizePrimitive(element);
}

export function normalizeIconDefinition(value: unknown): IconDefinition {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("The icon definition is invalid.");
  const record = value as Record<string, unknown>;
  const id = typeof record.id === "string" ? record.id.trim() : "";
  const name = typeof record.name === "string" ? record.name.trim() : "";
  const category = typeof record.category === "string" ? record.category.trim() : "";
  if (!/^[a-z0-9-]{1,48}$/.test(id) || !name || name.length > 80 || !category || category.length > 80) {
    throw new Error("The icon metadata is invalid.");
  }
  if (!Array.isArray(record.tags) || !record.tags.length || record.tags.length > 24) {
    throw new Error("The icon must have searchable tags.");
  }
  const tags = record.tags.map((tag) => {
    if (typeof tag !== "string" || !tag.trim() || tag.length > 80) throw new Error("The icon contains an invalid tag.");
    return tag.trim();
  });
  if (!Array.isArray(record.aliases) || record.aliases.length < 2 || record.aliases.length > 12) {
    throw new Error("The icon must have searchable aliases.");
  }
  const aliases = record.aliases.map((alias) => {
    if (typeof alias !== "string" || !alias.trim() || alias.length > 100) throw new Error("The icon contains an invalid alias.");
    return alias.trim();
  });
  if (!Array.isArray(record.primitives) || record.primitives.length < 2 || record.primitives.length > 64) {
    throw new Error("The icon must contain between 2 and 64 vector primitives.");
  }
  const primitives = record.primitives.map(normalizePrimitive);
  const rawElements = Array.isArray(record.elements) && record.elements.length ? record.elements : record.primitives;
  if (rawElements.length > 64) throw new Error("The icon contains too many vector elements.");
  const elements = rawElements.map(normalizeElement);
  return { id, name, category, aliases, tags, primitives, elements };
}
