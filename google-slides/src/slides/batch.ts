import type { Box } from "../core/geometry";

type PageElement = GoogleAppsScript.Slides.PageElement;

function dimensionPoints(dimension: GoogleAppsScript.Slides.Schema.Dimension | undefined): number | null {
  if (dimension?.magnitude == null) return null;
  if (dimension.unit === "EMU") return dimension.magnitude / 12700;
  return dimension.magnitude;
}

function collectApiElements(presentation: GoogleAppsScript.Slides.Schema.Presentation): Map<string, GoogleAppsScript.Slides.Schema.PageElement> {
  const result = new Map<string, GoogleAppsScript.Slides.Schema.PageElement>();
  const visit = (element: GoogleAppsScript.Slides.Schema.PageElement): void => {
    if (element.objectId) result.set(element.objectId, element);
    element.elementGroup?.children?.forEach(visit);
  };
  presentation.slides?.forEach((slide) => slide.pageElements?.forEach(visit));
  return result;
}

function boxRequest(
  element: PageElement,
  box: Box,
  apiElement: GoogleAppsScript.Slides.Schema.PageElement,
): GoogleAppsScript.Slides.Schema.Request | null {
  if (Math.abs(element.getRotation()) > 0.001) return null;
  const inherentWidth = dimensionPoints(apiElement.size?.width);
  const inherentHeight = dimensionPoints(apiElement.size?.height);
  if (!inherentWidth || !inherentHeight) return null;
  return {
    updatePageElementTransform: {
      objectId: element.getObjectId(),
      applyMode: "ABSOLUTE",
      transform: {
        scaleX: box.width / inherentWidth,
        scaleY: box.height / inherentHeight,
        shearX: 0,
        shearY: 0,
        translateX: box.left,
        translateY: box.top,
        unit: "PT",
      },
    },
  };
}

export function applyBoxesAtomically(
  presentationId: string,
  elements: PageElement[],
  boxes: Box[],
): boolean {
  if (!elements.length || elements.length !== boxes.length) return false;
  try {
    if (!Slides) return false;
    const presentation = Slides.Presentations.get(presentationId);
    const apiElements = collectApiElements(presentation);
    const byId = new Map(boxes.map((box) => [box.id, box]));
    const requests: GoogleAppsScript.Slides.Schema.Request[] = [];
    for (const element of elements) {
      const box = byId.get(element.getObjectId());
      const apiElement = apiElements.get(element.getObjectId());
      if (!box || !apiElement) return false;
      const request = boxRequest(element, box, apiElement);
      if (!request) return false;
      requests.push(request);
    }
    Slides.Presentations.batchUpdate({ requests, writeControl: { requiredRevisionId: presentation.revisionId } }, presentationId);
    return true;
  } catch (error) {
    console.warn(`Atomic Slides update failed; using SlidesApp fallback: ${String(error)}`);
    return false;
  }
}
