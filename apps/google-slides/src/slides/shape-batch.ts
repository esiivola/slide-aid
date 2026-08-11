type Request = GoogleAppsScript.Slides.Schema.Request;

function rgbColor(hex: string): GoogleAppsScript.Slides.Schema.RgbColor {
  const clean = hex.replace("#", "");
  return {
    red: Number.parseInt(clean.slice(0, 2), 16) / 255,
    green: Number.parseInt(clean.slice(2, 4), 16) / 255,
    blue: Number.parseInt(clean.slice(4, 6), 16) / 255,
  };
}

function elementProperties(pageObjectId: string, left: number, top: number, width: number, height: number): GoogleAppsScript.Slides.Schema.PageElementProperties {
  return {
    pageObjectId,
    size: { width: { magnitude: Math.max(0.1, width), unit: "PT" }, height: { magnitude: Math.max(0.1, height), unit: "PT" } },
    transform: { scaleX: 1, scaleY: 1, shearX: 0, shearY: 0, translateX: left, translateY: top, unit: "PT" },
  };
}

function objectId(): string {
  return `sa_${Utilities.getUuid().replace(/-/g, "").slice(0, 28)}`;
}

/**
 * Accumulates one atomic Slides `batchUpdate`. Both Chart Aid builders and icon
 * insertion use it: a chart or an icon is many small objects that must appear
 * together or not at all, and one round trip is the difference between a snappy
 * insert and dozens of sequential SlidesApp calls.
 */
export class ShapeBatch {
  private readonly requests: Request[] = [];
  private readonly objectIds: string[] = [];

  constructor(private readonly presentationId: string, private readonly slideId: string) {}

  get size(): number {
    return this.objectIds.length;
  }

  addShape(shapeType: string, left: number, top: number, width: number, height: number, color: string): string {
    const id = objectId();
    this.objectIds.push(id);
    this.requests.push(
      { createShape: { objectId: id, shapeType, elementProperties: elementProperties(this.slideId, left, top, width, height) } },
      {
        updateShapeProperties: {
          objectId: id,
          fields: "shapeBackgroundFill.solidFill,outline.propertyState",
          shapeProperties: {
            shapeBackgroundFill: { solidFill: { color: { rgbColor: rgbColor(color) }, alpha: 1 } },
            outline: { propertyState: "NOT_RENDERED" },
          },
        },
      },
    );
    return id;
  }

  /** Hollow shape: no fill, visible stroke. Used by outline-style icon rects and ellipses. */
  addOutlinedShape(shapeType: string, left: number, top: number, width: number, height: number, color: string, weight: number): string {
    const id = objectId();
    this.objectIds.push(id);
    this.requests.push(
      { createShape: { objectId: id, shapeType, elementProperties: elementProperties(this.slideId, left, top, width, height) } },
      {
        updateShapeProperties: {
          objectId: id,
          fields: "shapeBackgroundFill.propertyState,outline",
          shapeProperties: {
            shapeBackgroundFill: { propertyState: "NOT_RENDERED" },
            outline: {
              outlineFill: { solidFill: { color: { rgbColor: rgbColor(color) }, alpha: 1 } },
              weight: { magnitude: weight, unit: "PT" },
              dashStyle: "SOLID",
            },
          },
        },
      },
    );
    return id;
  }

  /**
   * Filled shape rotated by `radians` about the midpoint of its own top edge.
   * Slides has no adjustable pie geometry, so Harvey balls are built by fanning
   * rotated slivers out from the circle's center - this is the one place that
   * needs a transform the plain helpers cannot express.
   */
  addPivotedShape(shapeType: string, pivotX: number, pivotY: number, width: number, height: number, radians: number, color: string): string {
    const id = objectId();
    this.objectIds.push(id);
    const cos = Math.cos(radians);
    const sin = Math.sin(radians);
    this.requests.push(
      {
        createShape: {
          objectId: id,
          shapeType,
          elementProperties: {
            pageObjectId: this.slideId,
            size: { width: { magnitude: Math.max(0.1, width), unit: "PT" }, height: { magnitude: Math.max(0.1, height), unit: "PT" } },
            transform: {
              scaleX: cos,
              scaleY: cos,
              shearX: -sin,
              shearY: sin,
              translateX: pivotX - (cos * width) / 2,
              translateY: pivotY - (sin * width) / 2,
              unit: "PT",
            },
          },
        },
      },
      {
        updateShapeProperties: {
          objectId: id,
          fields: "shapeBackgroundFill.solidFill,outline.propertyState",
          shapeProperties: {
            shapeBackgroundFill: { solidFill: { color: { rgbColor: rgbColor(color) }, alpha: 1 } },
            outline: { propertyState: "NOT_RENDERED" },
          },
        },
      },
    );
    return id;
  }

  addLine(x1: number, y1: number, x2: number, y2: number, color: string, weight: number, dashed = false): string {
    const id = objectId();
    const width = Math.max(0.1, Math.abs(x2 - x1));
    const height = Math.max(0.1, Math.abs(y2 - y1));
    const properties = elementProperties(this.slideId, x1, y1, width, height);
    // A Slides line runs from its own (0,0) to (width,height); the transform
    // anchors that at (x1,y1). Negative scale mirrors it so leftward or upward
    // segments still end at (x2,y2) instead of pointing the wrong way.
    properties.transform = {
      ...properties.transform,
      scaleX: x2 >= x1 ? 1 : -1,
      scaleY: y2 >= y1 ? 1 : -1,
    };
    this.objectIds.push(id);
    this.requests.push(
      { createLine: { objectId: id, lineCategory: "STRAIGHT", elementProperties: properties } },
      {
        updateLineProperties: {
          objectId: id,
          fields: "lineFill.solidFill,weight,dashStyle",
          lineProperties: {
            lineFill: { solidFill: { color: { rgbColor: rgbColor(color) }, alpha: 1 } },
            weight: { magnitude: weight, unit: "PT" },
            dashStyle: dashed ? "DASH" : "SOLID",
          },
        },
      },
    );
    return id;
  }

  addText(text: string, left: number, top: number, width: number, height: number, size: number, alignment: string, color: string, bold = false): string {
    const id = objectId();
    this.objectIds.push(id);
    this.requests.push(
      { createShape: { objectId: id, shapeType: "TEXT_BOX", elementProperties: elementProperties(this.slideId, left, top, width, height) } },
      { insertText: { objectId: id, insertionIndex: 0, text } },
      {
        updateTextStyle: {
          objectId: id,
          textRange: { type: "ALL" },
          fields: "fontSize,foregroundColor,bold",
          style: { fontSize: { magnitude: size, unit: "PT" }, foregroundColor: { opaqueColor: { rgbColor: rgbColor(color) } }, bold },
        },
      },
      {
        updateParagraphStyle: {
          objectId: id,
          textRange: { type: "ALL" },
          fields: "alignment",
          style: { alignment },
        },
      },
    );
    return id;
  }

  /**
   * Attaches alt text to one already-queued object. Chart builders use it to
   * stamp each bar with the datum it represents, which is what lets the
   * annotation tools read real values off a clicked bar.
   */
  note(objectId: string, title: string, description: string): void {
    this.requests.push({ updatePageElementAltText: { objectId, title, description } });
  }

  commit(title: string, description: string): string {
    if (!this.objectIds.length) throw new Error("Nothing to insert.");
    if (!Slides) throw new Error("The Advanced Slides service is not enabled for this deployment.");
    const finalId = this.objectIds.length === 1 ? this.objectIds[0]! : objectId();
    if (this.objectIds.length > 1) this.requests.push({ groupObjects: { groupObjectId: finalId, childrenObjectIds: this.objectIds } });
    this.requests.push({ updatePageElementAltText: { objectId: finalId, title, description } });
    const presentation = Slides.Presentations.get(this.presentationId);
    Slides.Presentations.batchUpdate({ requests: this.requests, writeControl: { requiredRevisionId: presentation.revisionId } }, this.presentationId);
    return finalId;
  }
}
