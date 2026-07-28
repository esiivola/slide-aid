import { extractGoogleFileId } from "../core/integrations";
import type { ChartData, SheetSource } from "../core/chart-data";

export function readSheetData(spreadsheetUrl: string, sheetName: string, rangeA1: string): { data: ChartData; source: SheetSource } {
  const spreadsheetId = extractGoogleFileId(spreadsheetUrl);
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = sheetName.trim() ? spreadsheet.getSheetByName(sheetName.trim()) : spreadsheet.getSheets()[0];
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
  const cleanRange = rangeA1.trim();
  if (!cleanRange) throw new Error("Enter a source range such as A1:D6.");
  const values = sheet.getRange(cleanRange).getDisplayValues();
  return {
    data: { cells: values },
    source: {
      spreadsheetId,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
      sheetName: sheet.getName(),
      rangeA1: cleanRange,
    },
  };
}

export function refreshSheetData(source: SheetSource): ChartData {
  return readSheetData(source.spreadsheetUrl, source.sheetName, source.rangeA1).data;
}
