import { activeContext, elementBox, resolvePinnedReference } from "../slides/selection";
import { getDeckSettings, updateDeckSettings, type SavedFormat } from "../storage/document-state";

type PageElement = GoogleAppsScript.Slides.PageElement;
type Shape = GoogleAppsScript.Slides.Shape;

/**
 * The eight fixed palette colors from modColors.bas. PowerPoint offers these as
 * one-click swatches next to the theme colors; the sidebar shows the same names
 * so a deck built on either platform lands on the same greys and blues.
 */
export const GENERIC_PALETTE: readonly { name: string; hex: string }[] = [
  { name: "Dark blue", hex: "#1F497D" },
  { name: "Blue", hex: "#4F81BD" },
  { name: "Green", hex: "#9BBB59" },
  { name: "Red", hex: "#C0504D" },
  { name: "Orange", hex: "#F79646" },
  { name: "Purple", hex: "#8064A2" },
  { name: "Dark grey", hex: "#595959" },
  { name: "Light grey", hex: "#D9D9D9" },
];

function shapesIn(elements: readonly PageElement[]): Shape[] {
  const shapes: Shape[] = [];
  const visit = (list: readonly PageElement[]): void => {
    for (const element of list) {
      if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) visit(element.asGroup().getChildren());
      else if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) shapes.push(element.asShape());
    }
  };
  visit(elements);
  return shapes;
}

function solidHexOf(element: PageElement): string | null {
  if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return null;
  const fill = element.asShape().getFill();
  if (fill.getType() !== SlidesApp.FillType.SOLID) return null;
  const color = fill.getSolidFill().getColor();
  if (color.getColorType() === SlidesApp.ColorType.RGB) return color.asRgbColor().asHexString().toUpperCase();
  return null;
}

function themeNameOf(element: PageElement): string | null {
  if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return null;
  const fill = element.asShape().getFill();
  if (fill.getType() !== SlidesApp.FillType.SOLID) return null;
  const color = fill.getSolidFill().getColor();
  if (color.getColorType() === SlidesApp.ColorType.THEME) return String(color.asThemeColor().getThemeColorType());
  return null;
}

/** Reads everything Format Painter and My Formats copy from one shape. */
function readFormat(shape: Shape, name = ""): SavedFormat {
  const format: SavedFormat = { name, createdAt: new Date().toISOString() };
  const fill = shape.getFill();
  if (fill.getType() === SlidesApp.FillType.SOLID) {
    const color = fill.getSolidFill().getColor();
    if (color.getColorType() === SlidesApp.ColorType.RGB) format.fillColor = color.asRgbColor().asHexString();
  }
  const border = shape.getBorder();
  const lineFill = border.getLineFill();
  if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) {
    const color = lineFill.getSolidFill().getColor();
    if (color.getColorType() === SlidesApp.ColorType.RGB) format.borderColor = color.asRgbColor().asHexString();
  }
  const weight = border.getWeight();
  if (weight) format.borderWeight = weight;
  const dash = border.getDashStyle();
  if (dash) format.borderDash = String(dash);
  const style = shape.getText().getTextStyle();
  const family = style.getFontFamily();
  if (family) format.fontFamily = family;
  const size = style.getFontSize();
  if (size) format.fontSize = size;
  const foreground = style.getForegroundColor();
  if (foreground?.getColorType() === SlidesApp.ColorType.RGB) format.fontColor = foreground.asRgbColor().asHexString();
  const bold = style.isBold();
  if (bold != null) format.bold = bold;
  const italic = style.isItalic();
  if (italic != null) format.italic = italic;
  const alignment = shape.getText().getParagraphStyle().getParagraphAlignment();
  if (alignment) format.alignment = String(alignment);
  const contentAlignment = shape.getContentAlignment();
  if (contentAlignment) format.contentAlignment = String(contentAlignment);
  return format;
}

function writeFormat(shape: Shape, format: SavedFormat): void {
  if (format.fillColor) shape.getFill().setSolidFill(format.fillColor);
  const border = shape.getBorder();
  if (format.borderColor) border.getLineFill().setSolidFill(format.borderColor);
  if (format.borderWeight) border.setWeight(format.borderWeight);
  if (format.borderDash) {
    const dash = (SlidesApp.DashStyle as unknown as Record<string, GoogleAppsScript.Slides.DashStyle>)[format.borderDash];
    if (dash) border.setDashStyle(dash);
  }
  const text = shape.getText();
  if (text.asString().length) {
    const style = text.getTextStyle();
    if (format.fontFamily) style.setFontFamily(format.fontFamily);
    if (format.fontSize) style.setFontSize(format.fontSize);
    if (format.fontColor) style.setForegroundColor(format.fontColor);
    if (format.bold != null) style.setBold(format.bold);
    if (format.italic != null) style.setItalic(format.italic);
    if (format.alignment) {
      const alignment = (SlidesApp.ParagraphAlignment as unknown as Record<string, GoogleAppsScript.Slides.ParagraphAlignment>)[format.alignment];
      if (alignment) text.getParagraphStyle().setParagraphAlignment(alignment);
    }
  }
  if (format.contentAlignment) {
    const contentAlignment = (SlidesApp.ContentAlignment as unknown as Record<string, GoogleAppsScript.Slides.ContentAlignment>)[format.contentAlignment];
    if (contentAlignment) shape.setContentAlignment(contentAlignment);
  }
}

/** Format Painter: copy the pinned reference's whole look onto the other shapes. */
export function paintFormat(): { ok: true; message: string } {
  const context = activeContext(1);
  const reference = resolvePinnedReference(context);
  if (reference.getPageElementType() !== SlidesApp.PageElementType.SHAPE) throw new Error("Pin a shape as the reference before painting its format.");
  const format = readFormat(reference.asShape());
  const targets = shapesIn(context.elements.filter((element) => element.getObjectId() !== reference.getObjectId()));
  if (!targets.length) throw new Error("Select the shapes to paint; the pinned reference is excluded automatically.");
  targets.forEach((shape) => writeFormat(shape, format));
  return { ok: true, message: `Painted the reference format onto ${targets.length} shape${targets.length === 1 ? "" : "s"}.` };
}

/** Pick from Master: copy fill, line and/or font color from the pinned reference. */
export function pickColorsFromReference(target: "ALL" | "F" | "L" | "T"): { ok: true; message: string } {
  const context = activeContext(1);
  const reference = resolvePinnedReference(context);
  if (reference.getPageElementType() !== SlidesApp.PageElementType.SHAPE) throw new Error("Pin a shape as the reference first.");
  const source = readFormat(reference.asShape());
  const targets = shapesIn(context.elements.filter((element) => element.getObjectId() !== reference.getObjectId()));
  if (!targets.length) throw new Error("Select the objects to recolor; the pinned reference is excluded automatically.");
  for (const shape of targets) {
    if ((target === "ALL" || target === "F") && source.fillColor) shape.getFill().setSolidFill(source.fillColor);
    if ((target === "ALL" || target === "L") && source.borderColor) shape.getBorder().getLineFill().setSolidFill(source.borderColor);
    if ((target === "ALL" || target === "T") && source.fontColor && shape.getText().asString().length) {
      shape.getText().getTextStyle().setForegroundColor(source.fontColor);
    }
  }
  return { ok: true, message: `Copied the reference's colors to ${targets.length} object${targets.length === 1 ? "" : "s"}.` };
}

/** Color Info: report the selected object's fill as RGB, hex and theme link. */
export function colorInfo(): { message: string; hex: string | null; rgb: string | null; theme: string | null } {
  const context = activeContext(1);
  const element = context.elements[0]!;
  const hex = solidHexOf(element);
  const theme = themeNameOf(element);
  if (!hex && !theme) return { message: "The selected object has no solid fill.", hex: null, rgb: null, theme: null };
  if (theme) return { message: `Theme-linked fill: ${theme}.`, hex: null, rgb: null, theme };
  const channels = [1, 3, 5].map((start) => Number.parseInt(hex!.slice(start, start + 2), 16));
  const rgb = channels.join(", ");
  return { message: `Fill ${hex} · RGB ${rgb} · fixed (not theme-linked).`, hex, rgb, theme: null };
}

/** Select Similar: report the shapes on this slide matching the reference. */
export function selectSimilar(mode: "T" | "F" | "TF"): { ok: true; message: string; objectIds: string[] } {
  const context = activeContext(1);
  const reference = resolvePinnedReference(context);
  const referenceType = reference.getPageElementType() === SlidesApp.PageElementType.SHAPE ? String(reference.asShape().getShapeType()) : null;
  const referenceFill = solidHexOf(reference);
  const matches = context.slide.getPageElements().filter((element) => {
    if (element.getObjectId() === reference.getObjectId()) return false;
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return false;
    const sameType = referenceType !== null && String(element.asShape().getShapeType()) === referenceType;
    const sameFill = referenceFill !== null && solidHexOf(element) === referenceFill;
    if (mode === "T") return sameType;
    if (mode === "F") return sameFill;
    return sameType && sameFill;
  });
  if (!matches.length) throw new Error("No other shape on this slide matches the reference.");
  // Slides can only select a run of elements by re-selecting them one by one.
  matches[0]!.select();
  matches.slice(1).forEach((element) => element.select(false));
  return { ok: true, message: `Selected ${matches.length} similar shape${matches.length === 1 ? "" : "s"}.`, objectIds: matches.map((element) => element.getObjectId()) };
}

/** Fit to Text: shrink or grow each shape to the size of its text. */
export function fitToText(): { ok: true; message: string } {
  const context = activeContext(1);
  const shapes = shapesIn(context.elements).filter((shape) => shape.getText().asString().trim().length);
  if (!shapes.length) throw new Error("Select one or more shapes containing text.");
  const presentation = context.presentation;
  if (!Slides) throw new Error("The Advanced Slides service is not enabled for this deployment.");
  // SHAPE_AUTOFIT is the Slides equivalent of "resize shape to fit text". The
  // typings for ShapeProperties predate the autofit field, hence the cast.
  const requests = shapes.map((shape) => ({
    updateShapeProperties: {
      objectId: shape.getObjectId(),
      fields: "autofit.autofitType",
      shapeProperties: { autofit: { autofitType: "SHAPE_AUTOFIT" } },
    },
  })) as unknown as GoogleAppsScript.Slides.Schema.Request[];
  Slides.Presentations.batchUpdate({ requests }, presentation.getId());
  return { ok: true, message: `Fitted ${shapes.length} shape${shapes.length === 1 ? "" : "s"} to their text.` };
}

/**
 * Split at Cursor: the original keeps the text before the insertion point and a
 * copy below it takes the rest. Slides reports the caret as a text range, which
 * is the closest equivalent to PowerPoint's cursor position.
 */
export function splitAtCursor(): { ok: true; message: string } {
  const context = activeContext(1);
  const presentation = context.presentation;
  const selection = presentation.getSelection();
  const textRange = selection?.getTextRange();
  if (!textRange) throw new Error("Click into a text box at the point you want to split.");
  const element = context.elements[0];
  if (!element || element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) throw new Error("Click into a text box first.");
  const shape = element.asShape();
  const full = shape.getText().asString();
  const index = textRange.getStartIndex();
  if (index <= 0 || index >= full.length) throw new Error("Place the cursor inside the text, away from both ends.");
  const box = elementBox(element);
  const copy = shape.duplicate().asShape();
  shape.getText().setText(full.slice(0, index).replace(/\s+$/, ""));
  copy.getText().setText(full.slice(index).replace(/^\s+/, ""));
  copy.setLeft(box.left).setTop(box.top + box.height + 4).setWidth(box.width);
  copy.select();
  return { ok: true, message: "Split the text box; the second part sits directly below." };
}

/** Merge Boxes: every selected box becomes a paragraph of the first, in spatial order. */
export function mergeTextBoxes(): { ok: true; message: string } {
  const context = activeContext(2);
  const shapes = context.elements
    .filter((element) => element.getPageElementType() === SlidesApp.PageElementType.SHAPE)
    .sort((a, b) => (a.getTop() - b.getTop()) || (a.getLeft() - b.getLeft()));
  if (shapes.length < 2) throw new Error("Select at least two text boxes.");
  const target = shapes[0]!.asShape();
  const merged = shapes.map((element) => element.asShape().getText().asString().replace(/\s+$/, "")).filter((value) => value.length).join("\n");
  target.getText().setText(merged);
  shapes.slice(1).forEach((element) => element.remove());
  target.select();
  return { ok: true, message: `Merged ${shapes.length} text boxes into one.` };
}

/**
 * Snap to Table: move each object into the table cell it roughly sits over,
 * centred or aligned to one side with a margin.
 */
export function snapToTable(mode: "C" | "L" | "R", marginPt = 4): { ok: true; message: string } {
  const context = activeContext(2);
  const tableElement = context.elements.find((element) => element.getPageElementType() === SlidesApp.PageElementType.TABLE);
  if (!tableElement) throw new Error("Select the table together with the objects to snap.");
  const table = tableElement.asTable();
  const tableBox = elementBox(tableElement);

  const columnEdges: number[] = [tableBox.left];
  for (let column = 0; column < table.getNumColumns(); column += 1) {
    columnEdges.push(columnEdges[column]! + table.getColumn(column).getWidth());
  }
  const rowEdges: number[] = [tableBox.top];
  for (let row = 0; row < table.getNumRows(); row += 1) {
    rowEdges.push(rowEdges[row]! + table.getRow(row).getMinimumHeight());
  }

  const targets = context.elements.filter((element) => element.getObjectId() !== tableElement.getObjectId());
  if (!targets.length) throw new Error("Select the objects to snap as well as the table.");
  let moved = 0;
  for (const element of targets) {
    const box = elementBox(element);
    const centreX = box.left + box.width / 2;
    const centreY = box.top + box.height / 2;
    const column = columnEdges.findIndex((edge, index) => index < columnEdges.length - 1 && centreX >= edge && centreX < columnEdges[index + 1]!);
    const row = rowEdges.findIndex((edge, index) => index < rowEdges.length - 1 && centreY >= edge && centreY < rowEdges[index + 1]!);
    if (column < 0 || row < 0) continue;
    const cellLeft = columnEdges[column]!;
    const cellRight = columnEdges[column + 1]!;
    const cellTop = rowEdges[row]!;
    const cellBottom = rowEdges[row + 1]!;
    if (mode === "C") element.setLeft(cellLeft + (cellRight - cellLeft - box.width) / 2);
    else if (mode === "L") element.setLeft(cellLeft + marginPt);
    else element.setLeft(cellRight - box.width - marginPt);
    element.setTop(cellTop + (cellBottom - cellTop - box.height) / 2);
    moved += 1;
  }
  if (!moved) throw new Error("None of the selected objects sit over a table cell.");
  return { ok: true, message: `Snapped ${moved} object${moved === 1 ? "" : "s"} into their cells.` };
}

// ---------------------------------------------------------------------------
// My Formats - the reusable format library. PowerPoint keeps these in its
// sandbox folder; here they live in Document Properties so a deck's
// collaborators inherit the same set, like the named layouts do.
// ---------------------------------------------------------------------------

export function saveFormat(name: string): { ok: true; message: string } {
  const trimmed = String(name ?? "").trim();
  if (!trimmed) throw new Error("Give the format a name.");
  if (trimmed.length > 60) throw new Error("Format names are limited to 60 characters.");
  const context = activeContext(1);
  const source = context.elements.find((element) => element.getPageElementType() === SlidesApp.PageElementType.SHAPE);
  if (!source) throw new Error("Select a shape whose format should be saved.");
  const format = readFormat(source.asShape(), trimmed);
  const formats = getDeckSettings().formats.filter((entry) => entry.name !== trimmed);
  formats.push(format);
  formats.sort((a, b) => a.name.localeCompare(b.name));
  updateDeckSettings({ formats });
  return { ok: true, message: `Saved the format “${trimmed}”.` };
}

export function applyFormat(name: string): { ok: true; message: string } {
  const format = getDeckSettings().formats.find((entry) => entry.name === name);
  if (!format) throw new Error(`No saved format called “${name}”.`);
  const context = activeContext(1);
  const targets = shapesIn(context.elements);
  if (!targets.length) throw new Error("Select the shapes to format.");
  targets.forEach((shape) => writeFormat(shape, format));
  return { ok: true, message: `Applied “${format.name}” to ${targets.length} shape${targets.length === 1 ? "" : "s"}.` };
}

export function deleteFormat(name: string): { ok: true; message: string } {
  const formats = getDeckSettings().formats;
  const next = formats.filter((entry) => entry.name !== name);
  if (next.length === formats.length) throw new Error(`No saved format called “${name}”.`);
  updateDeckSettings({ formats: next });
  return { ok: true, message: `Deleted the format “${name}”.` };
}

export function listFormats(): SavedFormat[] {
  return getDeckSettings().formats;
}
