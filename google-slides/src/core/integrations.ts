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
