import { executeCommand, type CommandRequest } from "../commands/geometry-commands";
import {
  buildChart, buildLinkedChart, editChartData, rebuildChart, recolorSeries, refreshLinkedChart,
  restyleCharts, selectedChartState, setPalette, validateLinkedChart,
} from "../charts/charts";
import { deleteReference, getSettings, PALETTES } from "../storage/preferences";
import { activeContext, elementBox, labelFor, referenceState, setReferenceFromCurrentSelection } from "../slides/selection";
import { bounds } from "../core/geometry";
import { getDeckSettings } from "../storage/document-state";
import { applyPaletteToTheme, applyThemeColor, convertSelectionColors, currentThemeSwatches } from "../slides/theme";
import { applyLayout, deleteLayout, listLayouts, saveLayout } from "../layouts/layouts";
import { addSelectionToLibrary, configureLibrary, insertLibraryItem, listLibraryItems, refreshSelectedLibraryItem } from "../library/library";
import { fixQaIssue, focusQaIssue, scanDeck, type QaIssue } from "../qa/qa";

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
      theme: currentThemeSwatches(),
      layouts: listLayouts(),
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

export function scanSlideAidDeck(): ApiResponse {
  return response(scanDeck);
}

export function focusSlideAidQaIssue(slideId: string, objectId: string): ApiResponse {
  return response(() => focusQaIssue(slideId, objectId));
}

export function fixSlideAidQaIssue(issue: Pick<QaIssue, "type" | "slideId" | "objectIds">): ApiResponse {
  return response(() => fixQaIssue(issue));
}
