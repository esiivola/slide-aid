import { decodeLibraryReference, encodeLibraryReference, extractGoogleFileId } from "../core/integrations";
import { getDeckSettings, updateDeckSettings } from "../storage/document-state";
import { activeContext, elementBox } from "../slides/selection";

const MARKER_PREFIX = "Slide Aid library item:";

export interface LibraryItem {
  slideId: string;
  name: string;
  index: number;
  updatedAt?: string;
}

function libraryPresentation(): GoogleAppsScript.Slides.Presentation {
  const id = getDeckSettings().libraryPresentationId;
  if (!id) throw new Error("Configure a shared library presentation first.");
  return SlidesApp.openById(id);
}

export function configureLibrary(url: string): { message: string } {
  const id = extractGoogleFileId(url);
  const presentation = SlidesApp.openById(id);
  updateDeckSettings({ libraryPresentationId: id, libraryPresentationUrl: presentation.getUrl() });
  return { message: `Connected library “${presentation.getName()}”.` };
}

function itemName(slide: GoogleAppsScript.Slides.Slide, index: number): string {
  const marker = slide.getPageElements().find((element) => element.getTitle().startsWith(MARKER_PREFIX));
  if (marker) return marker.getTitle().slice(MARKER_PREFIX.length).trim();
  const text = slide.getShapes().map((shape) => shape.getText().asString().trim()).find(Boolean);
  return text?.slice(0, 80) || `Library item ${index + 1}`;
}

export function listLibraryItems(): LibraryItem[] {
  return libraryPresentation().getSlides().map((slide, index) => {
    const marker = slide.getPageElements().find((element) => element.getTitle().startsWith(MARKER_PREFIX));
    const updatedAt = marker?.getDescription().match(/Updated ([0-9T:.Z-]+)/)?.[1];
    return { slideId: slide.getObjectId(), name: itemName(slide, index), index, updatedAt };
  });
}

function insertFromSource(sourcePresentation: GoogleAppsScript.Slides.Presentation, slideId: string): GoogleAppsScript.Slides.PageElement {
  const source = sourcePresentation.getSlideById(slideId);
  if (!source) throw new Error("The library item no longer exists.");
  const context = activeContext();
  const inserted = source.getPageElements()
    .filter((element) => !element.getTitle().startsWith(MARKER_PREFIX))
    .map((element) => context.slide.insertPageElement(element));
  if (!inserted.length) throw new Error("The library item contains no insertable elements.");
  const result = (inserted.length === 1 ? inserted[0]! : context.slide.group(inserted)) as unknown as GoogleAppsScript.Slides.PageElement;
  const description = result.getDescription().trim();
  result.setDescription(`${description ? `${description}\n` : ""}Shared Slide Aid library component. ${encodeLibraryReference({ presentationId: sourcePresentation.getId(), slideId })}`);
  result.setTitle(`Shared library: ${itemName(source, 0)}`);
  result.select();
  return result;
}

export function insertLibraryItem(slideId: string): { message: string } {
  const library = libraryPresentation();
  insertFromSource(library, slideId);
  return { message: "Inserted a linked component from the shared library." };
}

export function refreshSelectedLibraryItem(): { message: string } {
  const context = activeContext(1);
  if (context.elements.length !== 1) throw new Error("Select exactly one linked library component.");
  const oldElement = context.elements[0]!;
  const reference = decodeLibraryReference(oldElement.getDescription());
  if (!reference) throw new Error("The selected object is not linked to a Slide Aid library item.");
  const oldBox = elementBox(oldElement);
  const sourcePresentation = SlidesApp.openById(reference.presentationId);
  const replacement = insertFromSource(sourcePresentation, reference.slideId);
  replacement.setWidth(oldBox.width).setHeight(oldBox.height).setLeft(oldBox.left).setTop(oldBox.top);
  oldElement.remove();
  replacement.select();
  return { message: "Refreshed the selected component from its shared library source." };
}

export function addSelectionToLibrary(name: string): { message: string } {
  const cleanName = name.trim();
  if (!cleanName) throw new Error("Enter a library item name.");
  const context = activeContext(1);
  const library = libraryPresentation();
  let slide = library.getSlides().find((candidate, index) => itemName(candidate, index).toLowerCase() === cleanName.toLowerCase());
  if (!slide) slide = library.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  else slide.getPageElements().filter((element) => !element.getTitle().startsWith(MARKER_PREFIX)).forEach((element) => element.remove());
  context.elements.forEach((element) => slide.insertPageElement(element));
  let marker = slide.getPageElements().find((element) => element.getTitle().startsWith(MARKER_PREFIX));
  if (!marker) {
    const markerShape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, -20, -20, 1, 1);
    markerShape.getFill().setTransparent();
    markerShape.getBorder().setTransparent();
    marker = markerShape as unknown as GoogleAppsScript.Slides.PageElement;
  }
  marker.setTitle(`${MARKER_PREFIX} ${cleanName}`);
  marker.setDescription(`Metadata marker for the Slide Aid shared element library. Updated ${new Date().toISOString()}.`);
  return { message: `Saved “${cleanName}” in the shared library. Existing inserted copies can now be refreshed.` };
}
