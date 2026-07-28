export interface Box {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
}

export type Axis = "H" | "V";
export type Edge = "L" | "R" | "T" | "B" | "CH" | "CV";

export const right = (box: Box): number => box.left + box.width;
export const bottom = (box: Box): number => box.top + box.height;
export const centerX = (box: Box): number => box.left + box.width / 2;
export const centerY = (box: Box): number => box.top + box.height / 2;

export function bounds(boxes: readonly Box[]): Box {
  if (boxes.length === 0) throw new Error("At least one object is required.");
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const r = Math.max(...boxes.map(right));
  const b = Math.max(...boxes.map(bottom));
  return { id: "selection-bounds", left, top, width: r - left, height: b - top };
}

export function sortSpatially(boxes: readonly Box[], axis: Axis): Box[] {
  return [...boxes].sort((a, b) => {
    const primary = axis === "H" ? a.left - b.left : a.top - b.top;
    const secondary = axis === "H" ? a.top - b.top : a.left - b.left;
    return primary || secondary || a.id.localeCompare(b.id);
  });
}

export function align(boxes: readonly Box[], reference: Box, edge: Edge): Box[] {
  return boxes.map((source) => {
    const box = { ...source };
    if (edge === "L") box.left = reference.left;
    if (edge === "R") box.left = right(reference) - box.width;
    if (edge === "T") box.top = reference.top;
    if (edge === "B") box.top = bottom(reference) - box.height;
    if (edge === "CH") box.left = centerX(reference) - box.width / 2;
    if (edge === "CV") box.top = centerY(reference) - box.height / 2;
    return box;
  });
}

export function dock(boxes: readonly Box[], reference: Box, direction: Exclude<Edge, "CH" | "CV">): Box[] {
  return boxes.map((source) => {
    const box = { ...source };
    if (direction === "L") box.left = right(reference);
    if (direction === "R") box.left = reference.left - box.width;
    if (direction === "T") box.top = bottom(reference);
    if (direction === "B") box.top = reference.top - box.height;
    return box;
  });
}

export function stretch(boxes: readonly Box[], reference: Box, direction: Exclude<Edge, "CH" | "CV">): Box[] {
  return boxes.map((source) => {
    const box = { ...source };
    if (direction === "L" && reference.left < right(box)) {
      box.width = right(box) - reference.left;
      box.left = reference.left;
    }
    if (direction === "R" && right(reference) > box.left) box.width = right(reference) - box.left;
    if (direction === "T" && reference.top < bottom(box)) {
      box.height = bottom(box) - reference.top;
      box.top = reference.top;
    }
    if (direction === "B" && bottom(reference) > box.top) box.height = bottom(reference) - box.top;
    return box;
  });
}

export function fillGap(boxes: readonly Box[], reference: Box, direction: Exclude<Edge, "CH" | "CV">): Box[] {
  return boxes.map((source) => {
    const box = { ...source };
    if (direction === "L" && box.left > right(reference)) {
      box.width = right(box) - right(reference);
      box.left = right(reference);
    }
    if (direction === "R" && right(box) < reference.left) box.width = reference.left - box.left;
    if (direction === "T" && box.top > bottom(reference)) {
      box.height = bottom(box) - bottom(reference);
      box.top = bottom(reference);
    }
    if (direction === "B" && bottom(box) < reference.top) box.height = reference.top - box.top;
    return box;
  });
}

export function matchSize(boxes: readonly Box[], reference: Box, dimension: "W" | "H" | "WH"): Box[] {
  return boxes.map((source) => {
    const cx = centerX(source);
    const cy = centerY(source);
    const box = { ...source };
    if (dimension.includes("W")) box.width = reference.width;
    if (dimension.includes("H")) box.height = reference.height;
    box.left = cx - box.width / 2;
    box.top = cy - box.height / 2;
    return box;
  });
}

export function stack(boxes: readonly Box[], axis: Axis, gap = 0): Box[] {
  const ordered = sortSpatially(boxes, axis);
  if (ordered.length === 0) return [];
  let x = ordered[0]!.left;
  let y = ordered[0]!.top;
  return ordered.map((source) => {
    const box = { ...source, left: x, top: y };
    if (axis === "H") x += source.width + gap;
    else y += source.height + gap;
    return box;
  });
}

export function setSpacing(boxes: readonly Box[], axis: Axis, gap: number): Box[] {
  return stack(boxes, axis, gap);
}

export function distribute(boxes: readonly Box[], axis: Axis): Box[] {
  if (boxes.length < 3) throw new Error("Distribute needs at least three objects.");
  const ordered = sortSpatially(boxes, axis);
  const start = axis === "H" ? Math.min(...boxes.map((box) => box.left)) : Math.min(...boxes.map((box) => box.top));
  const end = axis === "H" ? Math.max(...boxes.map(right)) : Math.max(...boxes.map(bottom));
  const total = boxes.reduce((sum, box) => sum + (axis === "H" ? box.width : box.height), 0);
  const gap = (end - start - total) / (boxes.length - 1);
  let position = start;
  return ordered.map((source) => {
    const box = { ...source };
    if (axis === "H") {
      box.left = position;
      position += box.width + gap;
    } else {
      box.top = position;
      position += box.height + gap;
    }
    return box;
  });
}

export function matrix(boxes: readonly Box[], columns: number, horizontalGap = 0, verticalGap = 0): Box[] {
  if (!Number.isInteger(columns) || columns < 1) throw new Error("Columns must be a positive integer.");
  const ordered = sortSpatially(boxes, "H");
  if (ordered.length === 0) return [];
  const width = Math.max(...ordered.map((box) => box.width));
  const height = Math.max(...ordered.map((box) => box.height));
  const left = ordered[0]!.left;
  const top = ordered[0]!.top;
  return ordered.map((source, index) => ({
    ...source,
    left: left + (index % columns) * (width + horizontalGap),
    top: top + Math.floor(index / columns) * (height + verticalGap),
  }));
}

export function scaleAroundCenter(boxes: readonly Box[], factor: number): Box[] {
  if (!(factor > 0)) throw new Error("Scale must be greater than zero.");
  return boxes.map((source) => {
    const width = source.width * factor;
    const height = source.height * factor;
    return {
      ...source,
      width,
      height,
      left: centerX(source) - width / 2,
      top: centerY(source) - height / 2,
    };
  });
}

export function placeRegion(slide: Box, preset: string, margin = 12): Box {
  let left = 0;
  let top = 0;
  let width = slide.width;
  let height = slide.height;
  if (preset === "LH" || preset === "RH") width /= 2;
  if (preset === "RH") left = slide.width / 2;
  if (preset === "TH" || preset === "BH") height /= 2;
  if (preset === "BH") top = slide.height / 2;
  if (["L3", "C3", "R3"].includes(preset)) width /= 3;
  if (preset === "C3") left = slide.width / 3;
  if (preset === "R3") left = (slide.width * 2) / 3;
  if (["Q1", "Q2", "Q3", "Q4"].includes(preset)) {
    width /= 2;
    height /= 2;
  }
  if (["Q2", "Q4"].includes(preset)) left = slide.width / 2;
  if (["Q3", "Q4"].includes(preset)) top = slide.height / 2;
  return {
    id: `region-${preset}`,
    left: left + margin,
    top: top + margin,
    width: Math.max(1, width - margin * 2),
    height: Math.max(1, height - margin * 2),
  };
}
