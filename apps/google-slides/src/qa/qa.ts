import { contrastRatio, isOutsideSlide } from "../core/integrations";
import { distribute, type Box } from "../core/geometry";
import { DATASHEET_PREFIX } from "../core/chart-data";
import { loadChartMetadata } from "../storage/document-state";
import { applyBoxesAtomically } from "../slides/batch";
import { elementBox } from "../slides/selection";
import { refreshSheetData } from "../integrations/sheets";

type PageElement = GoogleAppsScript.Slides.PageElement;

export type QaIssueType = "OFF_SLIDE" | "TINY_FONT" | "MISSING_ALT" | "LOW_CONTRAST" | "FIXED_RGB" | "ORPHAN_DATASHEET" | "STALE_LINK" | "BROKEN_LINK" | "IRREGULAR_SPACING";

export interface QaIssue {
  type: QaIssueType;
  severity: "error" | "warning" | "info";
  slideId: string;
  slideNumber: number;
  objectIds: string[];
  message: string;
  fixable: boolean;
}

function colorHex(color: GoogleAppsScript.Slides.Color, scheme: GoogleAppsScript.Slides.ColorScheme): string {
  return color.getColorType() === SlidesApp.ColorType.THEME
    ? scheme.getConcreteColor(color.asThemeColor().getThemeColorType()).asRgbColor().asHexString()
    : color.asRgbColor().asHexString();
}

function chartMetadata(element: PageElement) {
  return loadChartMetadata(element.getDescription());
}

function spacingIssue(slideId: string, slideNumber: number, elements: PageElement[]): QaIssue | null {
  if (elements.length < 3) return null;
  const ordered = elements.map((element) => ({ element, box: elementBox(element) })).sort((a, b) => a.box.left - b.box.left);
  const topSpan = Math.max(...ordered.map((item) => item.box.top)) - Math.min(...ordered.map((item) => item.box.top));
  if (topSpan > 8) return null;
  const gaps = ordered.slice(1).map((item, index) => item.box.left - (ordered[index]!.box.left + ordered[index]!.box.width));
  if (Math.max(...gaps) - Math.min(...gaps) <= 2) return null;
  return {
    type: "IRREGULAR_SPACING", severity: "info", slideId, slideNumber,
    objectIds: ordered.map((item) => item.element.getObjectId()),
    message: "A horizontal row has visibly inconsistent gaps.", fixable: true,
  };
}

export function scanDeck(): { issues: QaIssue[]; message: string } {
  const presentation = SlidesApp.getActivePresentation();
  const width = presentation.getPageWidth();
  const height = presentation.getPageHeight();
  const issues: QaIssue[] = [];
  const chartIds = new Set<string>();
  const datasheets: { slideId: string; slideNumber: number; element: PageElement; chartId: string }[] = [];

  presentation.getSlides().forEach((slide, slideIndex) => {
    const elements = slide.getPageElements();
    const scheme = slide.getColorScheme();
    elements.forEach((element) => {
      const id = element.getObjectId();
      const box = elementBox(element);
      if (isOutsideSlide(box, width, height) && !element.getTitle().startsWith("Slide Aid library item:")) {
        issues.push({ type: "OFF_SLIDE", severity: "warning", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: "Object extends beyond the slide canvas.", fixable: true });
      }
      if ([SlidesApp.PageElementType.IMAGE, SlidesApp.PageElementType.GROUP, SlidesApp.PageElementType.SHEETS_CHART].includes(element.getPageElementType()) && !element.getDescription().trim()) {
        issues.push({ type: "MISSING_ALT", severity: "warning", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: "Visual object has no alt-text description.", fixable: false });
      }
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const text = shape.getText().asString().trim();
        const size = text ? shape.getText().getTextStyle().getFontSize() : null;
        if (size != null && size < 12) issues.push({ type: "TINY_FONT", severity: "warning", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: `Text is ${size} pt.`, fixable: true });
        const fill = shape.getFill().getSolidFill();
        if (fill?.getColor().getColorType() === SlidesApp.ColorType.RGB) issues.push({ type: "FIXED_RGB", severity: "info", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: "Shape uses a fixed RGB fill instead of a theme-linked color.", fixable: false });
        const textColor = text ? shape.getText().getTextStyle().getForegroundColor() : null;
        if (text && fill && textColor) {
          const ratio = contrastRatio(colorHex(textColor, scheme), colorHex(fill.getColor(), scheme));
          if (ratio < 4.5) issues.push({ type: "LOW_CONTRAST", severity: "warning", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: `Text contrast is ${ratio.toFixed(1)}:1.`, fixable: false });
        }
      }
      if (element.getDescription().startsWith(DATASHEET_PREFIX)) datasheets.push({ slideId: slide.getObjectId(), slideNumber: slideIndex + 1, element, chartId: element.getDescription().slice(DATASHEET_PREFIX.length) });
      const metadata = chartMetadata(element);
      if (metadata) {
        chartIds.add(metadata.id);
        if (metadata.source) {
          try {
            const fresh = refreshSheetData(metadata.source);
            if (JSON.stringify(fresh.cells) !== JSON.stringify(metadata.data.cells)) issues.push({ type: "STALE_LINK", severity: "info", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: `Linked data changed in ${metadata.source.sheetName}!${metadata.source.rangeA1}. Select the chart and use Refresh selected.`, fixable: false });
          } catch {
            issues.push({ type: "BROKEN_LINK", severity: "error", slideId: slide.getObjectId(), slideNumber: slideIndex + 1, objectIds: [id], message: "Linked Google Sheets source cannot be read.", fixable: false });
          }
        }
      }
    });
    const spacing = spacingIssue(slide.getObjectId(), slideIndex + 1, elements.filter((element) => element.getPageElementType() !== SlidesApp.PageElementType.LINE));
    if (spacing) issues.push(spacing);
  });

  datasheets.filter((item) => !chartIds.has(item.chartId)).forEach((item) => issues.push({
    type: "ORPHAN_DATASHEET", severity: "info", slideId: item.slideId, slideNumber: item.slideNumber,
    objectIds: [item.element.getObjectId()], message: "Chart datasheet has no matching Slide Aid chart.", fixable: true,
  }));
  return { issues, message: `Found ${issues.length} QA issue${issues.length === 1 ? "" : "s"}.` };
}

export function focusQaIssue(slideId: string, objectId: string): { message: string } {
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSlideById(slideId);
  const element = presentation.getPageElementById(objectId);
  if (!slide || !element) throw new Error("The QA issue no longer exists.");
  slide.selectAsCurrentPage();
  element.select();
  return { message: "Selected the affected object." };
}

export function fixQaIssue(issue: Pick<QaIssue, "type" | "slideId" | "objectIds">): { message: string } {
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSlideById(issue.slideId);
  if (!slide) throw new Error("The affected slide no longer exists.");
  const elements = issue.objectIds.map((id) => presentation.getPageElementById(id)).filter((element): element is PageElement => Boolean(element));
  if (!elements.length) throw new Error("The affected objects no longer exist.");
  if (issue.type === "OFF_SLIDE") {
    const element = elements[0]!;
    const box = elementBox(element);
    element.setLeft(Math.min(Math.max(0, box.left), Math.max(0, presentation.getPageWidth() - box.width)));
    element.setTop(Math.min(Math.max(0, box.top), Math.max(0, presentation.getPageHeight() - box.height)));
  } else if (issue.type === "TINY_FONT") {
    elements.filter((element) => element.getPageElementType() === SlidesApp.PageElementType.SHAPE).forEach((element) => element.asShape().getText().getTextStyle().setFontSize(12));
  } else if (issue.type === "ORPHAN_DATASHEET") elements.forEach((element) => element.remove());
  else if (issue.type === "IRREGULAR_SPACING") {
    const boxes = elements.map(elementBox);
    const result = distribute(boxes, "H");
    if (!applyBoxesAtomically(presentation.getId(), elements, result)) result.forEach((box, index) => elements[index]!.setLeft(box.left));
  } else throw new Error("This issue requires human review and has no automatic fix.");
  return { message: `Fixed ${issue.type.toLowerCase().replace(/_/g, " ")}.` };
}
