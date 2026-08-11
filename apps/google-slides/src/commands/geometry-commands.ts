import {
  align, bounds, distribute, dock, fillGap, matchSize, matrix, placeRegion, processChain,
  scaleAroundCenter, setSpacing, squareColumns, stack, stretch, swapPositions,
  type Axis, type Box, type Edge, type SwapAnchor,
} from "../core/geometry";
import { getSettings, updateSettings } from "../storage/preferences";
import { applyBoxesAtomically } from "../slides/batch";
import { activeContext, elementBox, resolvePinnedReference, slideBox } from "../slides/selection";

const CM_TO_PT = 28.3464567;

export interface CommandRequest {
  command: string;
  argument?: string;
  referenceMode?: "PINNED" | "SLIDE" | "BOUNDS";
  gapCm?: number;
  columns?: number;
  rows?: number;
  scalePercent?: number;
  color?: string;
  colorTarget?: "F" | "L" | "T";
  /** Swap: which corner is matched, and whether sizes travel with positions. */
  swapAnchor?: SwapAnchor;
  swapSizes?: boolean;
  /** Magic Resizer scales font sizes with the geometry, like PowerPoint. */
  scaleFonts?: boolean;
}

export interface CommandResult {
  ok: true;
  message: string;
  selectedCount: number;
}

function targetsAndReference(request: CommandRequest): {
  context: ReturnType<typeof activeContext>;
  elements: GoogleAppsScript.Slides.PageElement[];
  boxes: Box[];
  reference: Box;
} {
  const context = activeContext(1);
  const mode = request.referenceMode ?? "PINNED";
  let elements = context.elements;
  let reference: Box;
  if (mode === "SLIDE") reference = slideBox(context);
  else if (mode === "BOUNDS") reference = bounds(elements.map(elementBox));
  else {
    const referenceElement = resolvePinnedReference(context);
    reference = elementBox(referenceElement);
    elements = elements.filter((element) => element.getObjectId() !== reference.id);
    if (elements.length === 0) throw new Error("Select one or more target objects; the pinned reference is excluded automatically.");
  }
  return { context, elements, boxes: elements.map(elementBox), reference };
}

function selectionOnly(minimum = 1): {
  context: ReturnType<typeof activeContext>;
  elements: GoogleAppsScript.Slides.PageElement[];
  boxes: Box[];
} {
  const context = activeContext(minimum);
  return { context, elements: context.elements, boxes: context.elements.map(elementBox) };
}

function applyBoxes(elements: GoogleAppsScript.Slides.PageElement[], boxes: Box[]): void {
  const presentation = SlidesApp.getActivePresentation();
  if (presentation && getSettings().useAtomicUpdates && applyBoxesAtomically(presentation.getId(), elements, boxes)) return;
  const byId = new Map(boxes.map((box) => [box.id, box]));
  for (const element of elements) {
    const box = byId.get(element.getObjectId());
    if (!box) continue;
    const old = elementBox(element);
    if (Math.abs(box.width - old.width) > 0.001) element.setWidth(Math.max(0.1, box.width));
    if (Math.abs(box.height - old.height) > 0.001) element.setHeight(Math.max(0.1, box.height));
    if (Math.abs(box.left - old.left) > 0.001) element.setLeft(box.left);
    if (Math.abs(box.top - old.top) > 0.001) element.setTop(box.top);
    if (box.rotation != null && Math.abs(box.rotation - old.rotation!) > 0.001) element.setRotation(box.rotation);
  }
}

function duplicateGrid(rows: number, columns: number, gap: number, slice: boolean): number {
  const { elements } = selectionOnly(1);
  if (elements.length !== 1) throw new Error("Select exactly one object.");
  if (rows < 1 || columns < 1 || !Number.isInteger(rows) || !Number.isInteger(columns)) throw new Error("Rows and columns must be positive integers.");
  const source = elements[0]!;
  const original = elementBox(source);
  const cellWidth = slice ? (original.width - (columns - 1) * gap) / columns : original.width;
  const cellHeight = slice ? (original.height - (rows - 1) * gap) / rows : original.height;
  if (cellWidth <= 0 || cellHeight <= 0) throw new Error("The gap is too large for the selected object.");
  if (slice) source.setWidth(cellWidth).setHeight(cellHeight);
  let count = 1;
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      if (row === 0 && column === 0) continue;
      const copy = source.duplicate();
      copy.setLeft(original.left + column * (cellWidth + gap));
      copy.setTop(original.top + row * (cellHeight + gap));
      count += 1;
    }
  }
  source.setLeft(original.left).setTop(original.top);
  return count;
}

function applyColor(request: CommandRequest): void {
  const { elements } = selectionOnly(1);
  const color = request.color;
  if (!color || !/^#[0-9a-f]{6}$/i.test(color)) throw new Error("Choose a valid color.");
  for (const element of elements) {
    const target = request.colorTarget ?? "F";
    if (target === "L") {
      if (element.getPageElementType() === SlidesApp.PageElementType.LINE) element.asLine().getLineFill().setSolidFill(color);
      else if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) element.asShape().getBorder().getLineFill().setSolidFill(color);
    } else if (target === "T" && element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      element.asShape().getText().getTextStyle().setForegroundColor(color);
    } else if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      element.asShape().getFill().setSolidFill(color);
    }
  }
}

function transformText(mode: string): void {
  const { elements } = selectionOnly(1);
  for (const element of elements) {
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
    const text = element.asShape().getText();
    const value = text.asString();
    let next = value;
    if (mode === "U") next = value.toUpperCase();
    if (mode === "L") next = value.toLowerCase();
    if (mode === "T") next = value.toLowerCase().replace(/(^|\s)\S/g, (part) => part.toUpperCase());
    if (mode === "S") next = value.toLowerCase().replace(/(^\s*|[.!?]\s+)\S/g, (part) => part.toUpperCase());
    if (mode === "TIDY") next = value.replace(/ {2,}/g, " ");
    text.setText(next);
  }
}

function swapText(): void {
  const { elements } = selectionOnly(2);
  if (elements.length !== 2 || elements.some((element) => element.getPageElementType() !== SlidesApp.PageElementType.SHAPE)) {
    throw new Error("Select exactly two shapes containing text.");
  }
  const first = elements[0]!.asShape().getText();
  const second = elements[1]!.asShape().getText();
  const value = first.asString();
  first.setText(second.asString());
  second.setText(value);
}

function alignAngles(request: CommandRequest): void {
  const { elements, boxes, reference } = targetsAndReference({ ...request, command: "angles" });
  applyBoxes(elements, boxes.map((box) => ({ ...box, rotation: reference.rotation ?? 0 })));
}

/**
 * Scales every text run in a shape. PowerPoint's Magic Resizer resizes type
 * along with geometry, which is the whole point of it - a scaled card whose text
 * stayed put looks broken.
 */
function scaleFontSizes(element: GoogleAppsScript.Slides.PageElement, factor: number): void {
  if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) {
    element.asGroup().getChildren().forEach((child) => scaleFontSizes(child, factor));
    return;
  }
  if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
  const text = element.asShape().getText();
  if (!text.asString().trim()) return;
  for (const run of text.getRuns()) {
    const style = run.getTextStyle();
    const size = style.getFontSize();
    if (size) style.setFontSize(Math.max(1, Math.round(size * factor * 10) / 10));
  }
}

export function executeCommand(request: CommandRequest): CommandResult {
  const command = request.command;
  let count = 0;
  if (["align", "dock", "stretch", "fillGap", "size", "golden"].includes(command)) {
    const { elements, boxes, reference } = targetsAndReference(request);
    let result = boxes;
    if (command === "align") result = align(boxes, reference, request.argument as Edge);
    if (command === "dock") result = dock(boxes, reference, request.argument as "L" | "R" | "T" | "B");
    if (command === "stretch") result = stretch(boxes, reference, request.argument as "L" | "R" | "T" | "B");
    if (command === "fillGap") result = fillGap(boxes, reference, request.argument as "L" | "R" | "T" | "B");
    if (command === "size") result = matchSize(boxes, reference, request.argument as "W" | "H" | "WH");
    if (command === "golden") result = boxes.map((box) => ({ ...box, top: reference.top + (reference.height - box.height) / 3 }));
    applyBoxes(elements, result);
    count = elements.length;
  } else if (command === "procChain") {
    const { elements, boxes, reference } = targetsAndReference(request);
    applyBoxes(elements, processChain(boxes, reference));
    count = elements.length;
  } else if (["stack", "spacing", "distribute", "matrix", "scale", "place", "swap"].includes(command)) {
    const { context, elements, boxes } = selectionOnly(command === "distribute" ? 3 : command === "swap" ? 2 : 1);
    let result = boxes;
    const gapCm = Number.isFinite(request.gapCm) ? request.gapCm! : getSettings().gapCm;
    const gap = gapCm * CM_TO_PT;
    if (command === "stack") result = stack(boxes, request.argument as Axis, gap);
    if (command === "spacing") result = setSpacing(boxes, request.argument as Axis, gap);
    if (command === "distribute") result = distribute(boxes, request.argument as Axis);
    if (command === "swap") result = swapPositions(boxes, request.swapAnchor ?? "C", request.swapSizes === true);
    if (command === "matrix") {
      // "Matrix" with no column count is PowerPoint's one-click near-square grid.
      const columns = request.columns && request.columns > 0 ? request.columns : squareColumns(boxes.length);
      result = matrix(boxes, columns, gap, gap);
      updateSettings({ gapCm, matrixColumns: columns });
    }
    if (command === "scale") {
      const factor = (request.scalePercent ?? 100) / 100;
      result = scaleAroundCenter(boxes, factor);
      if (request.scaleFonts !== false) elements.forEach((element) => scaleFontSizes(element, factor));
    }
    if (command === "place") {
      const region = placeRegion(slideBox(context), request.argument ?? "FULL");
      if (boxes.length === 1) result = [{ ...boxes[0]!, left: region.left, top: region.top, width: region.width, height: region.height }];
      else {
        const groupBounds = bounds(boxes);
        result = boxes.map((box) => ({ ...box, left: box.left + region.left - groupBounds.left, top: box.top + region.top - groupBounds.top }));
      }
    }
    if (command !== "matrix") updateSettings({ gapCm });
    applyBoxes(elements, result);
    count = elements.length;
  } else if (command === "slice" || command === "multiply") {
    count = duplicateGrid(request.rows ?? 2, request.columns ?? 2, (request.gapCm ?? 0.1) * CM_TO_PT, command === "slice");
  } else if (command === "color") {
    applyColor(request);
    count = activeContext().elements.length;
  } else if (command === "text") {
    transformText(request.argument ?? "TIDY");
    count = activeContext().elements.length;
  } else if (command === "swapText") {
    swapText();
    count = 2;
  } else if (command === "angles") {
    alignAngles(request);
    count = activeContext().elements.length;
  } else {
    throw new Error(`Unknown command: ${command}`);
  }
  return { ok: true, message: `Applied ${command} to ${count} object${count === 1 ? "" : "s"}.`, selectedCount: count };
}
