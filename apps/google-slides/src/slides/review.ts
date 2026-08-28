import { activeContext, elementBox, slideBox, type SelectionContext } from "./selection";
import { getSettings, updateSettings } from "../storage/preferences";
import { ShapeBatch } from "./shape-batch";

/**
 * On-slide review markup, modelled on the reference review toolbar (BCG "Cool
 * Macros" / PPT Productivity): coloured sticky notes, status stamps, and
 * callouts. Unlike Google Slides' native side-pane comments these are real
 * shapes, so they show in exported PDFs. Notes/stamps are plain (sharp) tagged
 * rectangles so the deck can be swept clean before the final send.
 */

const REVIEW_TAG = "slide-aid-review"; // sticky notes + callouts
const STAMP_TAG = "slide-aid-stamp";   // status stamps
const NOTE_WIDTH = 156;
const NOTE_HEIGHT = 50;
const CORNER_MARGIN = 10;
const CASCADE = 14;
const PAD = 5;
const NOTE_TEXT = "#222222";

// The six reference note colours; yellow is the default.
export type NoteColor = "YELLOW" | "GREEN" | "BLUE" | "PINK" | "LAVENDER" | "LAVBLUE";
const NOTE_COLORS: Record<NoteColor, string> = {
  YELLOW: "#FFE04F",
  GREEN: "#B2DF8A",
  BLUE: "#A0D2FF",
  PINK: "#FFB3C6",
  LAVENDER: "#D6C4F0",
  LAVBLUE: "#BEC8F5",
};
const LEADER_COLOR = "#E63C32";

// The eight reference status stamps.
export type StampKind = "DRAFT" | "WIP" | "NEW" | "UPDATED" | "CONFIDENTIAL" | "ONHOLD" | "OUTOFDATE" | "REMOVE";
const STAMP_LABELS: Record<StampKind, string> = {
  DRAFT: "DRAFT", WIP: "WORK IN PROGRESS", NEW: "NEW", UPDATED: "UPDATED",
  CONFIDENTIAL: "CONFIDENTIAL", ONHOLD: "ON HOLD", OUTOFDATE: "OUT OF DATE", REMOVE: "REMOVE",
};
const STAMP_COLORS: Record<StampKind, string> = {
  DRAFT: "#787878", WIP: "#E68C00", NEW: "#28A046", UPDATED: "#286EC8",
  CONFIDENTIAL: "#C81E1E", ONHOLD: "#8246A0", OUTOFDATE: "#C81E1E", REMOVE: "#C81E1E",
};

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function normalizeColor(value: string): NoteColor {
  const key = String(value).toUpperCase();
  return (key in NOTE_COLORS ? key : "YELLOW") as NoteColor;
}

function normalizeStamp(value: string): StampKind {
  const key = String(value).toUpperCase();
  return (key in STAMP_LABELS ? key : "DRAFT") as StampKind;
}

function description(element: GoogleAppsScript.Slides.PageElement): string {
  try {
    return element.getDescription() || "";
  } catch {
    return "";
  }
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

function stampLine(): string {
  const now = new Date();
  return `${reviewInitials()} · ${now.getDate()} ${MONTHS[now.getMonth()]}`;
}

function reviewNoteCount(context: SelectionContext): number {
  return context.slide.getPageElements().filter((element) => description(element).includes(`[${REVIEW_TAG}:`)).length;
}

function drawNote(batch: ShapeBatch, color: NoteColor, left: number, top: number, height: number, comment: string): void {
  batch.addShape("RECTANGLE", left, top, NOTE_WIDTH, height, NOTE_COLORS[color]);
  batch.addText(`${stampLine()}\n${comment}`, left + PAD, top + PAD, NOTE_WIDTH - 2 * PAD, height - 2 * PAD, 9, "START", NOTE_TEXT);
}

function noteHeight(comment: string): number {
  const lines = Math.max(1, Math.ceil(comment.length / 26));
  return Math.max(NOTE_HEIGHT, 20 + lines * 12);
}

export function addReviewNote(color: string, comment: string): { message: string; id: string } {
  const context = activeContext();
  const slide = slideBox(context);
  const col = normalizeColor(color);
  const height = noteHeight(comment);
  const left = slide.width - NOTE_WIDTH - CORNER_MARGIN;
  const top = CORNER_MARGIN + reviewNoteCount(context) * (NOTE_HEIGHT + CASCADE);
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  drawNote(batch, col, left, top, height, comment);
  const id = batch.commit("Slide Aid note", `[${REVIEW_TAG}:NOTE:${col}] ${comment}`);
  return { message: "Added a sticky note.", id };
}

export function addReviewCallout(comment: string): { message: string; id: string } {
  const context = activeContext(1);
  if (context.elements.length !== 1) throw new Error("Select exactly one object to point the callout at.");
  const target = elementBox(context.elements[0]!);
  const slide = slideBox(context);
  const height = noteHeight(comment);
  const left = Math.max(CORNER_MARGIN, Math.min(target.left + target.width + 16, slide.width - NOTE_WIDTH - CORNER_MARGIN));
  const top = Math.max(CORNER_MARGIN, target.top - height - 16);
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  drawNote(batch, "YELLOW", left, top, height, comment);
  batch.addLine(left + NOTE_WIDTH / 2, top + height, target.left + target.width / 2, target.top + target.height / 2, LEADER_COLOR, 2);
  const id = batch.commit("Slide Aid callout", `[${REVIEW_TAG}:CALLOUT] ${comment}`);
  return { message: "Added a callout.", id };
}

export function addStatusStamp(kind: string): { message: string; toggledOff: boolean; id?: string } {
  const context = activeContext();
  const k = normalizeStamp(kind);
  // Toggle off: if this stamp is already on the slide, remove it and stop.
  for (const element of context.slide.getPageElements()) {
    if (description(element).includes(`[${STAMP_TAG}:${k}]`)) {
      element.remove();
      return { message: `Removed the ${STAMP_LABELS[k]} stamp.`, toggledOff: true };
    }
  }
  const slide = slideBox(context);
  const label = STAMP_LABELS[k];
  const width = Math.max(140, label.length * 13 + 28);
  const height = 40;
  const left = (slide.width - width) / 2;
  const top = 18;
  const batch = new ShapeBatch(context.presentation.getId(), context.slide.getObjectId());
  batch.addShape("RECTANGLE", left, top, width, height, STAMP_COLORS[k]);
  batch.addText(label, left, top + 8, width, height - 16, 20, "CENTER", "#FFFFFF", true);
  const id = batch.commit(`Slide Aid stamp ${k}`, `[${STAMP_TAG}:${k}]`);
  return { message: `Stamped ${label}.`, toggledOff: false, id };
}

export function removeReviewMarkup(): { message: string; removed: number } {
  const presentation = SlidesApp.getActivePresentation();
  if (!presentation) throw new Error("No Google Slides presentation is active.");
  let removed = 0;
  for (const slide of presentation.getSlides()) {
    for (const element of slide.getPageElements()) {
      const desc = description(element);
      if (desc.includes(`[${REVIEW_TAG}`) || desc.includes(`[${STAMP_TAG}`)) {
        element.remove();
        removed += 1;
      }
    }
  }
  return { message: `Removed ${removed} review mark${removed === 1 ? "" : "s"} from the deck.`, removed };
}
