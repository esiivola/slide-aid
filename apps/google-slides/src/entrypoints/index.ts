import { executeCommand, type CommandRequest } from "../commands/geometry-commands";
import {
  applyChartSettings, buildChart, buildLinkedChart, chartSettingsState, editChartData, familyPalette,
  rebuildChart, recolorSeries, refreshLinkedChart, resetChartSettings, resetFamilyPalette, restyleCharts,
  saveFamilyPalette, selectedChartState, setPalette, validateLinkedChart,
} from "../charts/charts";
import {
  annotateAverageLine, annotateDifference, annotateValueLine, cycleElementState, insertCheckbox, insertHarveyBall,
} from "../charts/annotations";
import { dataLayouts, insertSampleSlides, removeSampleSlides } from "../charts/samples";
import { deleteReference, getSettings, PALETTES } from "../storage/preferences";
import { activeContext, elementBox, labelFor, referenceState, setReferenceFromCurrentSelection } from "../slides/selection";
import { bounds } from "../core/geometry";
import { getDeckSettings } from "../storage/document-state";
import { applyPaletteToTheme, applyThemeColor, convertSelectionColors, currentThemeSwatches } from "../slides/theme";
import { applyLayout, deleteLayout, listLayouts, saveLayout } from "../layouts/layouts";
import { addSelectionToLibrary, configureLibrary, insertLibraryItem, listLibraryItems, refreshSelectedLibraryItem } from "../library/library";
import { fixQaIssue, focusQaIssue, scanDeck, type QaIssue } from "../qa/qa";
import {
  applyFormat, colorInfo, deleteFormat, fitToText, GENERIC_PALETTE, listFormats, mergeTextBoxes,
  paintFormat, pickColorsFromReference, saveFormat, selectSimilar, snapToTable, splitAtCursor,
} from "../commands/object-commands";
import {
  buildAgenda, copyToAllSlides, hideObjects, removeGeneratedAgenda, removeSpeakerNotes, slideSummary, unhideAll,
} from "../commands/deck-commands";
import { insertCuratedIcon, insertEditableIcon, insertIconImage, makeIconsEditable } from "../slides/icons";
import { iconPathsFor } from "../slides/icon-catalog";

export interface ApiResponse<T = unknown> {
  ok: boolean;
  data?: T;
  message?: string;
  error?: string;
}

function response<T>(operation: () => T): ApiResponse<T> {
  try {
    const data = operation();
    const message = typeof data === "object" && data && "message" in data ? String((data as { message: unknown }).message) : undefined;
    return { ok: true, data, message };
  } catch (error) {
    console.error(error);
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

export function onOpen(_event?: GoogleAppsScript.Events.AppsScriptEvent): void {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem("Open Slide Aid", "showSlideAidSidebar")
    .addSeparator()
    .addItem("Build column chart", "menuBuildColumn")
    .addItem("Rebuild selected chart", "rebuildSlideAidChart")
    .addItem("Refresh linked chart", "refreshLinkedSlideAidChart")
    .addToUi();
}

// Kept global through the generated wrapper so a basic menu item can invoke it.
export function menuBuildColumn(): ApiResponse {
  return buildSlideAidChart("COL");
}

export function onInstall(event?: GoogleAppsScript.Events.AppsScriptEvent): void {
  onOpen(event);
}

export function showSlideAidSidebar(): void {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("Slide Aid");
  SlidesApp.getUi().showSidebar(html);
}

export function getSidebarState(): ApiResponse {
  return response(() => {
    const context = activeContext();
    const settings = getSettings();
    const deckSettings = getDeckSettings();
    return {
      presentationName: context.presentation.getName(),
      slideId: context.slide.getObjectId(),
      selectedCount: context.elements.length,
      selection: context.elements.map((element) => ({ label: labelFor(element), type: String(element.getPageElementType()), ...elementBox(element) })),
      selectionBounds: context.elements.length ? bounds(context.elements.map(elementBox)) : null,
      reference: referenceState(),
      settings: { ...settings, palette: deckSettings.palette ?? settings.palette },
      deckSettings,
      palettes: Object.keys(PALETTES),
      genericPalette: GENERIC_PALETTE,
      theme: currentThemeSwatches(),
      layouts: listLayouts(),
      formats: listFormats(),
      selectedChart: context.elements.length ? selectedChartState() : null,
    };
  });
}

export function setReferenceFromSelection(): ApiResponse {
  return response(() => ({ reference: setReferenceFromCurrentSelection(), message: "Reference pinned." }));
}

export function clearReference(): ApiResponse {
  return response(() => {
    deleteReference();
    return { message: "Reference cleared." };
  });
}

export function runSlideAidCommand(request: CommandRequest): ApiResponse {
  return response(() => executeCommand(request));
}

export function buildSlideAidChart(kind: string): ApiResponse {
  return response(() => buildChart(kind));
}

export function rebuildSlideAidChart(): ApiResponse {
  return response(rebuildChart);
}

export function editSlideAidChartData(): ApiResponse {
  return response(editChartData);
}

export function restyleSlideAidCharts(allCharts: boolean): ApiResponse {
  return response(() => restyleCharts(Boolean(allCharts)));
}

export function setSlideAidPalette(name: string): ApiResponse {
  return response(() => setPalette(name));
}

export function recolorSlideAidSeries(seriesIndex: number, color: string): ApiResponse {
  return response(() => recolorSeries(Number(seriesIndex), color));
}

export function buildLinkedSlideAidChart(kind: string, spreadsheetUrl: string, sheetName: string, rangeA1: string): ApiResponse {
  return response(() => buildLinkedChart(kind, spreadsheetUrl, sheetName, rangeA1));
}

export function validateLinkedSlideAidChart(kind: string, spreadsheetUrl: string, sheetName: string, rangeA1: string): ApiResponse {
  return response(() => validateLinkedChart(kind, spreadsheetUrl, sheetName, rangeA1));
}

export function refreshLinkedSlideAidChart(): ApiResponse {
  return response(refreshLinkedChart);
}

export function applySlideAidThemeColor(target: "F" | "L" | "T", themeName: string): ApiResponse {
  return response(() => applyThemeColor(target, themeName));
}

export function convertSlideAidColors(toTheme: boolean): ApiResponse {
  return response(() => convertSelectionColors(Boolean(toTheme)));
}

export function applySlideAidPaletteToTheme(paletteName: string): ApiResponse {
  return response(() => applyPaletteToTheme(paletteName));
}

export function saveSlideAidLayout(name: string): ApiResponse {
  return response(() => saveLayout(name));
}

export function applySlideAidLayout(name: string): ApiResponse {
  return response(() => applyLayout(name));
}

export function deleteSlideAidLayout(name: string): ApiResponse {
  return response(() => deleteLayout(name));
}

export function configureSlideAidLibrary(url: string): ApiResponse {
  return response(() => configureLibrary(url));
}

export function getSlideAidLibraryItems(): ApiResponse {
  return response(listLibraryItems);
}

export function insertSlideAidLibraryItem(slideId: string): ApiResponse {
  return response(() => insertLibraryItem(slideId));
}

export function addSelectionToSlideAidLibrary(name: string): ApiResponse {
  return response(() => addSelectionToLibrary(name));
}

export function refreshSelectedSlideAidLibraryItem(): ApiResponse {
  return response(refreshSelectedLibraryItem);
}

export function insertSlideAidIcon(id: string, name: string, color: string, pngBase64: string): ApiResponse {
  return response(() => insertIconImage(id, name, color, pngBase64));
}

export function insertSlideAidEditableIcon(id: string, name: string, color: string): ApiResponse {
  return response(() => insertEditableIcon(id, name, color));
}

export function insertSlideAidCuratedIcon(icon: unknown, color: string, strokeWidth?: number): ApiResponse {
  return response(() => insertCuratedIcon(icon, color, strokeWidth ?? 1.6));
}

export function makeSlideAidIconsEditable(): ApiResponse {
  return response(makeIconsEditable);
}

export function getSlideAidIconPaths(ids: string[]): ApiResponse {
  return response(() => ({ paths: iconPathsFor(ids) }));
}

export function scanSlideAidDeck(): ApiResponse {
  return response(scanDeck);
}

export function focusSlideAidQaIssue(slideId: string, objectId: string): ApiResponse {
  return response(() => focusQaIssue(slideId, objectId));
}

export function fixSlideAidQaIssue(issue: Pick<QaIssue, "type" | "slideId" | "objectIds">): ApiResponse {
  return response(() => fixQaIssue(issue));
}

// --- Chart Aid: Style ------------------------------------------------------

export function getSlideAidChartSettings(): ApiResponse {
  return response(chartSettingsState);
}

export function applySlideAidChartSettings(scope: string, values: Record<string, unknown>): ApiResponse {
  return response(() => applyChartSettings(scope, values));
}

export function resetSlideAidChartSettings(scope: string): ApiResponse {
  return response(() => resetChartSettings(scope));
}

export function getSlideAidFamilyPalette(family: string): ApiResponse {
  return response(() => familyPalette(family));
}

export function saveSlideAidFamilyPalette(family: string, colors: unknown): ApiResponse {
  return response(() => saveFamilyPalette(family, colors));
}

export function resetSlideAidFamilyPalette(family: string): ApiResponse {
  return response(() => resetFamilyPalette(family));
}

// --- Chart Aid: Data, Annotations, Elements --------------------------------

export function getSlideAidDataLayouts(): ApiResponse {
  return response(() => ({ layouts: dataLayouts() }));
}

export function insertSlideAidSampleSlides(): ApiResponse {
  return response(insertSampleSlides);
}

export function removeSlideAidSampleSlides(): ApiResponse {
  return response(removeSampleSlides);
}

export function annotateSlideAidDifference(mode: string, periods?: number): ApiResponse {
  return response(() => annotateDifference(mode === "PCT" ? "PCT" : mode === "CAGR" ? "CAGR" : "ABS", periods));
}

export function annotateSlideAidValueLine(value: number): ApiResponse {
  return response(() => annotateValueLine(Number(value)));
}

export function annotateSlideAidAverageLine(): ApiResponse {
  return response(annotateAverageLine);
}

export function insertSlideAidHarveyBall(percent: number, color: string): ApiResponse {
  return response(() => insertHarveyBall(Number(percent), color));
}

export function insertSlideAidCheckbox(state: string, color: string): ApiResponse {
  return response(() => insertCheckbox(state, color));
}

export function cycleSlideAidElementState(color: string): ApiResponse {
  return response(() => cycleElementState(color));
}

// --- Slide Aid: Wizards, Color, Text, Shape --------------------------------

export function paintSlideAidFormat(): ApiResponse {
  return response(paintFormat);
}

export function saveSlideAidFormat(name: string): ApiResponse {
  return response(() => saveFormat(name));
}

export function applySlideAidFormat(name: string): ApiResponse {
  return response(() => applyFormat(name));
}

export function deleteSlideAidFormat(name: string): ApiResponse {
  return response(() => deleteFormat(name));
}

export function selectSlideAidSimilar(mode: string): ApiResponse {
  return response(() => selectSimilar(mode === "F" ? "F" : mode === "TF" ? "TF" : "T"));
}

export function pickSlideAidColors(target: string): ApiResponse {
  return response(() => pickColorsFromReference(target === "F" || target === "L" || target === "T" ? target : "ALL"));
}

export function getSlideAidColorInfo(): ApiResponse {
  return response(colorInfo);
}

export function fitSlideAidShapesToText(): ApiResponse {
  return response(fitToText);
}

export function splitSlideAidTextAtCursor(): ApiResponse {
  return response(splitAtCursor);
}

export function mergeSlideAidTextBoxes(): ApiResponse {
  return response(mergeTextBoxes);
}

export function snapSlideAidToTable(mode: string, marginCm?: number): ApiResponse {
  const margin = Number.isFinite(marginCm) ? Number(marginCm) * 28.3464567 : 4;
  return response(() => snapToTable(mode === "L" ? "L" : mode === "R" ? "R" : "C", margin));
}

// --- Slide Aid: View & Expert ---------------------------------------------

export function hideSlideAidObjects(): ApiResponse {
  return response(hideObjects);
}

export function unhideSlideAidObjects(): ApiResponse {
  return response(unhideAll);
}

export function pasteSlideAidOnSlides(): ApiResponse {
  return response(copyToAllSlides);
}

export function removeSlideAidSpeakerNotes(): ApiResponse {
  return response(removeSpeakerNotes);
}

export function getSlideAidSlideSummary(): ApiResponse {
  return response(slideSummary);
}

export function buildSlideAidAgenda(items: unknown): ApiResponse {
  return response(() => buildAgenda(items));
}

export function removeSlideAidAgenda(): ApiResponse {
  return response(() => removeGeneratedAgenda());
}
