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

export class ChartBatch {
  private readonly requests: Request[] = [];
  private readonly objectIds: string[] = [];

  constructor(private readonly presentationId: string, private readonly slideId: string) {}

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

  addLine(x1: number, y1: number, x2: number, y2: number, color: string, weight: number): string {
    const id = objectId();
    const width = Math.max(0.1, Math.abs(x2 - x1));
    const height = Math.max(0.1, Math.abs(y2 - y1));
    const properties = elementProperties(this.slideId, x1, y1, width, height);
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
          fields: "lineFill.solidFill,weight",
          lineProperties: {
            lineFill: { solidFill: { color: { rgbColor: rgbColor(color) }, alpha: 1 } },
            weight: { magnitude: weight, unit: "PT" },
          },
        },
      },
    );
    return id;
  }

  addText(text: string, left: number, top: number, width: number, height: number, size: number, alignment: string, color: string): string {
    const id = objectId();
    this.objectIds.push(id);
    this.requests.push(
      { createShape: { objectId: id, shapeType: "TEXT_BOX", elementProperties: elementProperties(this.slideId, left, top, width, height) } },
      { insertText: { objectId: id, insertionIndex: 0, text } },
      {
        updateTextStyle: {
          objectId: id,
          textRange: { type: "ALL" },
          fields: "fontSize,foregroundColor",
          style: { fontSize: { magnitude: size, unit: "PT" }, foregroundColor: { opaqueColor: { rgbColor: rgbColor(color) } } },
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

  commit(title: string, description: string): string {
    if (!this.objectIds.length) throw new Error("Chart contains no drawable elements.");
    if (!Slides) throw new Error("The Advanced Slides service is not enabled for this deployment.");
    const finalId = this.objectIds.length === 1 ? this.objectIds[0]! : objectId();
    if (this.objectIds.length > 1) this.requests.push({ groupObjects: { groupObjectId: finalId, childrenObjectIds: this.objectIds } });
    this.requests.push({ updatePageElementAltText: { objectId: finalId, title, description } });
    const presentation = Slides.Presentations.get(this.presentationId);
    Slides.Presentations.batchUpdate({ requests: this.requests, writeControl: { requiredRevisionId: presentation.revisionId } }, this.presentationId);
    return finalId;
  }
}
