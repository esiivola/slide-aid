import type { Box } from "../core/geometry";
import { deleteReference, getReference, saveReference, type ReferenceRecord } from "../storage/preferences";

export interface SelectionContext {
  presentation: GoogleAppsScript.Slides.Presentation;
  slide: GoogleAppsScript.Slides.Slide;
  elements: GoogleAppsScript.Slides.PageElement[];
}

export function activeContext(minimum = 0): SelectionContext {
  const presentation = SlidesApp.getActivePresentation();
  if (!presentation) throw new Error("No Google Slides presentation is active.");
  const selection = presentation.getSelection();
  if (!selection) throw new Error("No active slide or selection was found.");
  const page = selection.getCurrentPage();
  if (!page || page.getPageType() !== SlidesApp.PageType.SLIDE) throw new Error("Open a normal slide before using Slide Aid.");
  const range = selection.getPageElementRange();
  const elements = range?.getPageElements() ?? [];
  if (elements.length < minimum) throw new Error(`Select at least ${minimum} object${minimum === 1 ? "" : "s"}.`);
  return { presentation, slide: page.asSlide(), elements };
}

export function elementBox(element: GoogleAppsScript.Slides.PageElement): Box {
  const width = element.getWidth();
  const height = element.getHeight();
  if (width == null || height == null) throw new Error("The selected object has no editable size.");
  return {
    id: element.getObjectId(),
    left: element.getLeft(),
    top: element.getTop(),
    width,
    height,
    rotation: element.getRotation(),
  };
}

export function slideBox(context: SelectionContext): Box {
  return {
    id: "slide",
    left: 0,
    top: 0,
    width: context.presentation.getPageWidth(),
    height: context.presentation.getPageHeight(),
  };
}

export function labelFor(element: GoogleAppsScript.Slides.PageElement): string {
  const title = element.getTitle().trim();
  if (title) return title;
  return `${String(element.getPageElementType())} (${element.getObjectId().slice(0, 8)})`;
}

export function setReferenceFromCurrentSelection(): ReferenceRecord {
  const context = activeContext(1);
  if (context.elements.length !== 1) throw new Error("Select exactly one object to set as the reference.");
  const element = context.elements[0]!;
  const reference: ReferenceRecord = {
    presentationId: context.presentation.getId(),
    slideId: context.slide.getObjectId(),
    objectId: element.getObjectId(),
    label: labelFor(element),
  };
  saveReference(reference);
  return reference;
}

export function resolvePinnedReference(context: SelectionContext): GoogleAppsScript.Slides.PageElement {
  const reference = getReference();
  if (!reference || reference.presentationId !== context.presentation.getId()) {
    throw new Error("No reference is pinned for this presentation. Select one object and click Set reference.");
  }
  if (reference.slideId !== context.slide.getObjectId()) throw new Error("The pinned reference is on another slide.");
  const element = context.presentation.getPageElementById(reference.objectId);
  if (!element) {
    deleteReference();
    throw new Error("The pinned reference no longer exists. Set a new reference.");
  }
  return element;
}

export function referenceState(): ReferenceRecord | null {
  const reference = getReference();
  if (!reference) return null;
  const presentation = SlidesApp.getActivePresentation();
  if (!presentation || presentation.getId() !== reference.presentationId) return null;
  if (!presentation.getPageElementById(reference.objectId)) {
    deleteReference();
    return null;
  }
  return reference;
}
