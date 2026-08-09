import { activeContext, elementBox } from "../slides/selection";

type PageElement = GoogleAppsScript.Slides.PageElement;

// Slides has no per-object visibility flag, so "hidden" objects are parked
// off-canvas with their real position recorded in alt text. That keeps the
// PowerPoint promise - position and layer order survive - using the one
// mechanism Slides gives us.
const HIDDEN_TAG = /\[slide-aid-hidden:(-?[\d.]+):(-?[\d.]+)\]/;
const PARK_OFFSET = 20000;

export function hideObjects(): { ok: true; message: string } {
  const context = activeContext(1);
  let hidden = 0;
  for (const element of context.elements) {
    if (HIDDEN_TAG.test(element.getDescription())) continue;
    const box = elementBox(element);
    const description = element.getDescription();
    element.setDescription(`${description} [slide-aid-hidden:${box.left}:${box.top}]`.trim());
    element.setLeft(box.left + PARK_OFFSET);
    hidden += 1;
  }
  if (!hidden) throw new Error("The selected objects are already hidden.");
  return { ok: true, message: `Hid ${hidden} object${hidden === 1 ? "" : "s"}. Use Unhide All to bring them back.` };
}

export function unhideAll(): { ok: true; message: string } {
  const context = activeContext();
  let restored = 0;
  for (const element of context.slide.getPageElements()) {
    const match = element.getDescription().match(HIDDEN_TAG);
    if (!match) continue;
    element.setLeft(Number(match[1]));
    element.setTop(Number(match[2]));
    element.setDescription(element.getDescription().replace(HIDDEN_TAG, "").trim());
    restored += 1;
  }
  if (!restored) throw new Error("No hidden objects on this slide.");
  return { ok: true, message: `Restored ${restored} object${restored === 1 ? "" : "s"}.` };
}

/** Paste on Slides: copy the selection onto every other slide at the same spot. */
export function copyToAllSlides(): { ok: true; message: string } {
  const context = activeContext(1);
  const sources = context.elements.map((element) => ({ element, box: elementBox(element) }));
  const slides = context.presentation.getSlides().filter((slide) => slide.getObjectId() !== context.slide.getObjectId());
  if (!slides.length) throw new Error("This presentation has only one slide.");
  for (const slide of slides) {
    for (const { element, box } of sources) {
      const copy = slide.insertPageElement(element as unknown as PageElement);
      copy.setLeft(box.left).setTop(box.top);
    }
  }
  return { ok: true, message: `Pasted ${sources.length} object${sources.length === 1 ? "" : "s"} onto ${slides.length} slide${slides.length === 1 ? "" : "s"}.` };
}

export function removeSpeakerNotes(): { ok: true; message: string } {
  const context = activeContext();
  let cleared = 0;
  for (const slide of context.presentation.getSlides()) {
    const notes = slide.getNotesPage().getSpeakerNotesShape();
    if (!notes || !notes.getText().asString().trim()) continue;
    notes.getText().setText("");
    cleared += 1;
  }
  return { ok: true, message: cleared ? `Cleared speaker notes on ${cleared} slide${cleared === 1 ? "" : "s"}.` : "No speaker notes to remove." };
}

/** Copy Slide Summary: the deck's slide titles as plain text for the clipboard. */
export function slideSummary(): { message: string; summary: string } {
  const context = activeContext();
  const lines = context.presentation.getSlides().map((slide, index) => {
    const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE) ?? slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
    const title = titlePlaceholder ? titlePlaceholder.asShape().getText().asString().trim().replace(/\s+/g, " ") : "";
    return `${index + 1}. ${title || "(untitled)"}`;
  });
  return { message: `Summarized ${lines.length} slide${lines.length === 1 ? "" : "s"}.`, summary: lines.join("\n") };
}

/**
 * Agenda: an overview slide plus a separator before each section, with the
 * current item highlighted. Google Slides has no section concept, so sections
 * are the deck's slide titles - re-running replaces the generated slides.
 */
const AGENDA_TAG = "[slide-aid-agenda]";

export function buildAgenda(items: unknown): { ok: true; message: string } {
  const context = activeContext();
  const presentation = context.presentation;
  const entries = Array.isArray(items)
    ? items.map((item) => String(item).trim()).filter((item) => item.length)
    : [];
  if (entries.length < 2) throw new Error("Enter at least two agenda items, one per line.");
  if (entries.length > 12) throw new Error("An agenda takes at most 12 items.");

  removeGeneratedAgenda(true);
  const width = presentation.getPageWidth();
  const height = presentation.getPageHeight();

  const addAgendaSlide = (highlight: number): void => {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    const heading = slide.insertTextBox("Agenda", 48, 40, width - 96, 34);
    heading.getText().getTextStyle().setFontSize(22).setBold(true);
    heading.setDescription(AGENDA_TAG);
    entries.forEach((entry, index) => {
      const line = slide.insertTextBox(`${index + 1}.  ${entry}`, 48, 96 + index * 30, width - 96, 28);
      const style = line.getText().getTextStyle();
      style.setFontSize(15);
      // The current item is the only one at full strength - the rest recede.
      if (highlight === index) style.setBold(true).setForegroundColor("#1F4E79");
      else style.setForegroundColor("#8A8F94");
      line.setDescription(AGENDA_TAG);
    });
    if (96 + entries.length * 30 > height) throw new Error("Too many agenda items to fit on a slide.");
  };

  addAgendaSlide(-1);
  entries.forEach((_, index) => addAgendaSlide(index));
  return { ok: true, message: `Generated ${entries.length + 1} agenda slides. Move them where you need them.` };
}

export function removeGeneratedAgenda(quiet = false): { ok: true; message: string } {
  const context = activeContext();
  let removed = 0;
  for (const slide of context.presentation.getSlides()) {
    if (!slide.getPageElements().some((element) => element.getDescription().includes(AGENDA_TAG))) continue;
    slide.remove();
    removed += 1;
  }
  if (!removed && !quiet) throw new Error("This presentation has no generated agenda slides.");
  return { ok: true, message: `Removed ${removed} agenda slide${removed === 1 ? "" : "s"}.` };
}
