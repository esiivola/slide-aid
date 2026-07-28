import { normalizeLayout, projectLayout } from "../core/integrations";
import { sortSpatially } from "../core/geometry";
import { getDeckSettings, updateDeckSettings, type LayoutPreset } from "../storage/document-state";
import { applyBoxesAtomically } from "../slides/batch";
import { activeContext, elementBox } from "../slides/selection";

function orderedSelection() {
  const context = activeContext(1);
  const boxes = sortSpatially(context.elements.map(elementBox), "V");
  const byId = new Map(context.elements.map((element) => [element.getObjectId(), element]));
  return { context, boxes, elements: boxes.map((box) => byId.get(box.id)!) };
}

export function listLayouts(): LayoutPreset[] {
  return getDeckSettings().layouts;
}

export function saveLayout(name: string): { message: string } {
  const cleanName = name.trim();
  if (!cleanName) throw new Error("Enter a layout name.");
  const { context, boxes } = orderedSelection();
  const preset: LayoutPreset = {
    name: cleanName,
    slots: normalizeLayout(boxes, context.presentation.getPageWidth(), context.presentation.getPageHeight()),
    createdAt: new Date().toISOString(),
  };
  const layouts = getDeckSettings().layouts.filter((layout) => layout.name.toLowerCase() !== cleanName.toLowerCase());
  layouts.push(preset);
  layouts.sort((a, b) => a.name.localeCompare(b.name));
  updateDeckSettings({ layouts });
  return { message: `Saved layout “${cleanName}” with ${boxes.length} slots.` };
}

export function applyLayout(name: string): { message: string } {
  const preset = getDeckSettings().layouts.find((layout) => layout.name === name);
  if (!preset) throw new Error(`Layout not found: ${name}`);
  const { context, elements, boxes } = orderedSelection();
  const result = projectLayout(preset.slots, boxes, context.presentation.getPageWidth(), context.presentation.getPageHeight());
  if (!applyBoxesAtomically(context.presentation.getId(), elements, result)) {
    result.forEach((box, index) => elements[index]!.setWidth(box.width).setHeight(box.height).setLeft(box.left).setTop(box.top));
  }
  return { message: `Applied layout “${name}”.` };
}

export function deleteLayout(name: string): { message: string } {
  const layouts = getDeckSettings().layouts.filter((layout) => layout.name !== name);
  updateDeckSettings({ layouts });
  return { message: `Deleted layout “${name}”.` };
}
