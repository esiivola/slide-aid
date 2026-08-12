import { flattenIconElements, flattenPath, polylineSegments, type IconPolyline } from "../core/icon-path";
import { fillSpans, isFilledIcon } from "../core/icon-fill";
import { normalizeIconDefinition } from "../core/integrations";
import { iconPaths } from "./icon-catalog";
import { activeContext, elementBox } from "./selection";
import { ShapeBatch } from "./shape-batch";

type PageElement = GoogleAppsScript.Slides.PageElement;

// Inserted edge length in points; one inch, matching the task pane's default.
const ICON_SIZE_PT = 72;
const DESIGN_GRID = 24;
const DEFAULT_STROKE = 1.6;

// Mirrors PowerPoint's "IconAid:<id>[:<#hex>]" shape tag, so both products
// recognise an inserted icon the same way.
const ICON_TAG = /\[iconaid:([a-z0-9-]{1,64})(?::(#[0-9a-fA-F]{6}))?\]/;

function iconTag(id: string, color: string): string {
  return `[iconaid:${id}:${color.toUpperCase()}]`;
}

function requireColor(color: string): string {
  if (!/^#[0-9a-f]{6}$/i.test(color)) throw new Error("Icon color must use #RRGGBB format.");
  return color.toUpperCase();
}

function requireIconId(id: string): string {
  const value = String(id);
  if (!/^[a-z0-9-]{1,64}$/.test(value)) throw new Error("The icon id is invalid.");
  return value;
}

function catalogRuns(id: string): IconPolyline[] | null {
  const subpaths = iconPaths(id);
  if (!subpaths?.length) return null;
  const runs: IconPolyline[] = [];
  // Some upstream icon packages repeat identical SVG paths (377 catalog
  // entries currently do this). Rendering those duplicates with even-odd fill
  // cancels the geometry completely, so normalize them once at the boundary.
  for (const subpath of new Set(subpaths)) {
    try {
      runs.push(...flattenPath(subpath));
    } catch {
      // The catalog is built from normalized paths, but one bad subpath should
      // not prevent the remaining geometry in the icon from being inserted.
    }
  }
  return runs.length ? runs : null;
}

function addCatalogGeometry(
  batch: ShapeBatch,
  id: string,
  color: string,
  box: { left: number; top: number; width: number; height: number },
  runs: readonly IconPolyline[],
): void {
  const scaleX = box.width / DESIGN_GRID;
  const scaleY = box.height / DESIGN_GRID;
  if (isFilledIcon(id)) {
    for (const span of fillSpans(runs)) {
      batch.addShape(
        "RECTANGLE",
        box.left + span.left * scaleX, box.top + span.top * scaleY,
        span.width * scaleX, span.height * scaleY, color,
      );
    }
    return;
  }

  const weight = DEFAULT_STROKE * Math.min(scaleX, scaleY);
  for (const segment of polylineSegments(runs)) {
    batch.addLine(
      box.left + segment.x1 * scaleX, box.top + segment.y1 * scaleY,
      box.left + segment.x2 * scaleX, box.top + segment.y2 * scaleY,
      color, weight,
    );
  }
}

/**
 * Inserts a legacy icon as a tagged picture.
 *
 * Kept for compatibility with older sidebar deployments and presentations.
 * Current sidebar insertions use insertEditableIcon instead.
 */
export function insertIconImage(id: string, name: string, color: string, pngBase64: string): { ok: true; message: string } {
  const iconId = requireIconId(id);
  const hex = requireColor(color);
  const payload = String(pngBase64 ?? "");
  if (!payload || payload.length > 4_000_000 || !/^[A-Za-z0-9+/=]+$/.test(payload)) throw new Error("The icon preview could not be read.");

  const context = activeContext();
  const left = (context.presentation.getPageWidth() - ICON_SIZE_PT) / 2;
  const top = (context.presentation.getPageHeight() - ICON_SIZE_PT) / 2;
  const blob = Utilities.newBlob(Utilities.base64Decode(payload), "image/png", `${iconId}.png`);
  const image = context.slide.insertImage(blob, left, top, ICON_SIZE_PT, ICON_SIZE_PT);
  image.setTitle(`IconAid: ${name || iconId}`);
  image.setDescription(`Slide Aid icon. Use Make Editable to turn it into shapes. ${iconTag(iconId, hex)}`);
  image.select();
  return { ok: true, message: `Inserted ${name || iconId}. Select it and click Make Editable to turn it into shapes.` };
}

/** Inserts any catalog icon immediately as grouped, editable Slides geometry. */
export function insertEditableIcon(id: string, name: string, color: string): { ok: true; message: string } {
  const iconId = requireIconId(id);
  const hex = requireColor(color);
  const runs = catalogRuns(iconId);
  if (!runs) throw new Error(`The icon data for ${iconId} is not in this deployment's catalog.`);

  const context = activeContext();
  const box = {
    left: (context.presentation.getPageWidth() - ICON_SIZE_PT) / 2,
    top: (context.presentation.getPageHeight() - ICON_SIZE_PT) / 2,
    width: ICON_SIZE_PT,
    height: ICON_SIZE_PT,
  };
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  addCatalogGeometry(batch, iconId, hex, box, runs);
  const label = String(name || iconId);
  const objectId = batch.commit(`IconAid: ${label}`, `Editable Slide Aid icon ${iconTag(iconId, hex)}`);
  const inserted = context.presentation.getPageElementById(objectId);
  if (inserted) inserted.select();
  return { ok: true, message: `Inserted ${label} as ${batch.size} editable shape${batch.size === 1 ? "" : "s"}.` };
}

/** Curated schema-3 icons carry their own primitives and insert as shapes directly. */
export function insertCuratedIcon(icon: unknown, color: string, strokeWidth = DEFAULT_STROKE): { ok: true; message: string } {
  const definition = normalizeIconDefinition(icon);
  const hex = requireColor(color);
  if (!Number.isFinite(strokeWidth) || strokeWidth <= 0 || strokeWidth > 4) throw new Error("Icon stroke width must be between 0 and 4.");

  const context = activeContext();
  const scale = ICON_SIZE_PT / DESIGN_GRID;
  const weight = strokeWidth * scale;
  const left = (context.presentation.getPageWidth() - ICON_SIZE_PT) / 2;
  const top = (context.presentation.getPageHeight() - ICON_SIZE_PT) / 2;
  const { shapes, polylines } = flattenIconElements(definition.elements);
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());

  for (const shape of shapes) {
    const shapeType = shape.kind === "ellipse" ? "ELLIPSE" : "RECTANGLE";
    const box = [left + shape.x * scale, top + shape.y * scale, shape.width * scale, shape.height * scale] as const;
    if (shape.filled) batch.addShape(shapeType, box[0], box[1], box[2], box[3], hex);
    else batch.addOutlinedShape(shapeType, box[0], box[1], box[2], box[3], hex, weight);
  }
  for (const segment of polylineSegments(polylines)) {
    batch.addLine(left + segment.x1 * scale, top + segment.y1 * scale, left + segment.x2 * scale, top + segment.y2 * scale, hex, weight);
  }
  const objectId = batch.commit(`IconAid: ${definition.name}`, `Editable Slide Aid icon ${iconTag(definition.id, hex)}`);
  const inserted = context.presentation.getPageElementById(objectId);
  if (inserted) inserted.select();
  return { ok: true, message: `Inserted ${definition.name} as ${batch.size} editable vector${batch.size === 1 ? "" : "s"}.` };
}

interface TaggedIcon {
  element: PageElement;
  id: string;
  color: string;
}

function taggedIcons(elements: readonly PageElement[]): TaggedIcon[] {
  const found: TaggedIcon[] = [];
  for (const element of elements) {
    const match = element.getDescription().match(ICON_TAG);
    if (!match) continue;
    // Only pictures need converting; an already-converted icon keeps its tag.
    if (element.getPageElementType() !== SlidesApp.PageElementType.IMAGE) continue;
    found.push({ element, id: match[1]!, color: match[2]?.toUpperCase() ?? "#1F497D" });
  }
  return found;
}

/**
 * Replaces inserted icon pictures with grouped native shapes, the Google Slides
 * counterpart of PowerPoint's Make Editable button.
 *
 * Stroke icons become line runs at the catalog's stroke weight; solid icons are
 * scan-converted into filled slices, because Slides cannot fill an arbitrary
 * outline. Either way the result is real, recolourable, ungroupable geometry
 * occupying the picture's exact box.
 */
export function makeIconsEditable(): { ok: true; message: string } {
  const context = activeContext();
  const selected = context.elements.length ? context.elements : context.slide.getPageElements();
  const icons = taggedIcons(selected);
  if (!icons.length) {
    throw new Error(context.elements.length
      ? "Select one or more inserted icons first."
      : "No inserted icons on this slide to make editable.");
  }

  let converted = 0;
  let objects = 0;
  const missing: string[] = [];
  for (const icon of icons) {
    const runs = catalogRuns(icon.id);
    if (!runs) {
      missing.push(icon.id);
      continue;
    }
    const box = elementBox(icon.element);
    const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
    // Honour the picture's own box so a resized legacy icon converts at the
    // size the user actually gave it.
    addCatalogGeometry(batch, icon.id, icon.color, box, runs);
    if (!batch.size) {
      missing.push(icon.id);
      continue;
    }
    batch.commit(icon.element.getTitle() || `IconAid: ${icon.id}`, `Editable Slide Aid icon ${iconTag(icon.id, icon.color)}`);
    objects += batch.size;
    icon.element.remove();
    converted += 1;
  }

  if (!converted) throw new Error(`The icon data for ${missing.slice(0, 3).join(", ")} is not in this deployment's catalog.`);
  const skipped = missing.length ? ` ${missing.length} could not be read.` : "";
  return { ok: true, message: `Converted ${converted} icon${converted === 1 ? "" : "s"} into ${objects} editable shapes.${skipped}` };
}
