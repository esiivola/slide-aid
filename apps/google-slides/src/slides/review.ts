import { activeContext, elementBox, slideBox, type SelectionContext } from "./selection";
import { getSettings, updateSettings } from "../storage/preferences";
import { ShapeBatch } from "./shape-batch";

/**
 * On-slide review markup: sticky comment notes, TODO/EDIT markers, and callouts
 * that point at an object. Unlike Google Slides' native side-pane comments these
 * are real shapes, so they show in exported PDFs and printouts - which is the
 * whole point of a review pass. Every mark carries a `[slide-aid-review:...]` tag
 * in its alt text so the deck can be swept clean before the final send.
 */

export type ReviewKind = "NOTE" | "TODO" | "EDIT";

const REVIEW_TAG = "slide-aid-review";
const NOTE_WIDTH = 156;
const NOTE_HEIGHT = 52;
const CORNER_MARGIN = 10;
const CASCADE = 16;
const PAD = 6;

// Deliberately loud, fixed colors - review marks must be impossible to miss, and
// they are removed before export, so matching the deck theme is a non-goal.
const STYLES: Record<ReviewKind, { fill: string; text: string; label: string }> = {
  NOTE: { fill: "#FFD60A", text: "#111111", label: "" },
  TODO: { fill: "#FF3B30", text: "#FFFFFF", label: "TODO" },
  EDIT: { fill: "#FF9F0A", text: "#111111", label: "EDIT" },
};
const LEADER_COLOR = "#FF3B30";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function isReviewMark(description: string | null | undefined): boolean {
  return !!description && description.includes(`[${REVIEW_TAG}`);
}

/** Compress a name or email local-part into up to three uppercase initials. */
export function initialsFrom(value: string): string {
  const parts = value.replace(/[._+-]+/g, " ").trim().split(/\s+/).filter(Boolean);
  return parts.map((part) => part[0]!.toUpperCase()).join("").slice(0, 3);
}

export function reviewInitials(): string {
  const stored = getSettings().initials?.trim();
  if (stored) return stored;
  let email = "";
  try {
    email = Session.getActiveUser().getEmail() || "";
  } catch {
    email = "";
  }
  const local = (email.split("@")[0] ?? "").trim();
  const derived = initialsFrom(local);
  // Single-token logins can't form real initials - show the whole name rather
  // than a lonely letter; the user can shorten it with the initials field.
  return derived.length >= 2 ? derived : local || "?";
}

export function setReviewInitials(value: string): { message: string; initials: string } {
  const clean = value.trim().slice(0, 6);
  updateSettings({ initials: clean });
  return { message: clean ? `Review initials set to ${clean}.` : "Review initials cleared; they will be derived from your account.", initials: clean };
}

function stamp(kind: ReviewKind): string {
  const now = new Date();
  const dated = `${reviewInitials()} · ${now.getDate()} ${MONTHS[now.getMonth()]}`;
  const label = STYLES[kind].label;
  return label ? `${label} · ${dated}` : dated;
}

function reviewCountOnSlide(context: SelectionContext): number {
  return context.slide.getPageElements().filter((element) => isReviewMark(element.getDescription())).length;
}

function placeNote(batch: ShapeBatch, kind: ReviewKind, left: number, top: number, comment: string): void {
  const style = STYLES[kind];
  const text = `${stamp(kind)}\n${comment}`.trim();
  batch.addShape("ROUND_RECTANGLE", left, top, NOTE_WIDTH, NOTE_HEIGHT, style.fill);
  batch.addText(text, left + PAD, top + PAD, NOTE_WIDTH - 2 * PAD, NOTE_HEIGHT - 2 * PAD, 9, "START", style.text, true);
}

export function addReviewNote(kind: ReviewKind, comment: string): { message: string; id: string } {
  const context = activeContext();
  const slide = slideBox(context);
  const n = reviewCountOnSlide(context);
  const left = slide.width - NOTE_WIDTH - CORNER_MARGIN;
  const top = CORNER_MARGIN + n * CASCADE;
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  placeNote(batch, kind, left, top, comment);
  const id = batch.commit(`Slide Aid ${kind}`, `[${REVIEW_TAG}:${kind}] ${comment}`);
  return { message: `Added a ${STYLES[kind].label || "comment"} mark.`, id };
}

export function addReviewCallout(comment: string): { message: string; id: string } {
  const context = activeContext(1);
  if (context.elements.length !== 1) throw new Error("Select exactly one object to point the callout at.");
  const target = elementBox(context.elements[0]!);
  const slide = slideBox(context);
  const left = Math.max(CORNER_MARGIN, Math.min(target.left + target.width + 16, slide.width - NOTE_WIDTH - CORNER_MARGIN));
  const top = Math.max(CORNER_MARGIN, target.top - NOTE_HEIGHT - 16);
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  placeNote(batch, "NOTE", left, top, comment);
  // Leader from the note's bottom-center to the target's center.
  batch.addLine(left + NOTE_WIDTH / 2, top + NOTE_HEIGHT, target.left + target.width / 2, target.top + target.height / 2, LEADER_COLOR, 2.25);
  const id = batch.commit("Slide Aid callout", `[${REVIEW_TAG}:CALLOUT] ${comment}`);
  return { message: "Added a callout.", id };
}

export function removeReviewMarkup(): { message: string; removed: number } {
  const presentation = SlidesApp.getActivePresentation();
  if (!presentation) throw new Error("No Google Slides presentation is active.");
  let removed = 0;
  for (const slide of presentation.getSlides()) {
    for (const element of slide.getPageElements()) {
      let description = "";
      try {
        description = element.getDescription();
      } catch {
        description = "";
      }
      if (isReviewMark(description)) {
        element.remove();
        removed += 1;
      }
    }
  }
  return { message: `Removed ${removed} review mark${removed === 1 ? "" : "s"} from the deck.`, removed };
}
