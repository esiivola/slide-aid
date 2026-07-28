import { PALETTES } from "../storage/preferences";
import { activeContext } from "./selection";

type PageElement = GoogleAppsScript.Slides.PageElement;
type ThemeColorType = GoogleAppsScript.Slides.ThemeColorType;

const THEME_NAMES = ["DARK1", "LIGHT1", "DARK2", "LIGHT2", "ACCENT1", "ACCENT2", "ACCENT3", "ACCENT4", "ACCENT5", "ACCENT6"] as const;

function themeMap(): Record<string, ThemeColorType> {
  return {
    DARK1: SlidesApp.ThemeColorType.DARK1,
    LIGHT1: SlidesApp.ThemeColorType.LIGHT1,
    DARK2: SlidesApp.ThemeColorType.DARK2,
    LIGHT2: SlidesApp.ThemeColorType.LIGHT2,
    ACCENT1: SlidesApp.ThemeColorType.ACCENT1,
    ACCENT2: SlidesApp.ThemeColorType.ACCENT2,
    ACCENT3: SlidesApp.ThemeColorType.ACCENT3,
    ACCENT4: SlidesApp.ThemeColorType.ACCENT4,
    ACCENT5: SlidesApp.ThemeColorType.ACCENT5,
    ACCENT6: SlidesApp.ThemeColorType.ACCENT6,
  };
}

function visit(elements: PageElement[], callback: (element: PageElement) => void): void {
  elements.forEach((element) => {
    if (element.getPageElementType() === SlidesApp.PageElementType.GROUP) visit(element.asGroup().getChildren(), callback);
    else callback(element);
  });
}

export interface ThemeSwatch {
  name: string;
  hex: string;
}

export function currentThemeSwatches(): ThemeSwatch[] {
  const context = activeContext();
  const scheme = context.slide.getColorScheme();
  const mapping = themeMap();
  return THEME_NAMES.map((name) => ({ name, hex: scheme.getConcreteColor(mapping[name]!).asRgbColor().asHexString() }));
}

export function applyThemeColor(target: "F" | "L" | "T", themeName: string): { message: string } {
  const context = activeContext(1);
  const theme = themeMap()[themeName];
  if (!theme) throw new Error(`Unknown theme color: ${themeName}`);
  visit(context.elements, (element) => {
    if (target === "T" && element.getPageElementType() === SlidesApp.PageElementType.SHAPE) element.asShape().getText().getTextStyle().setForegroundColor(theme);
    else if (target === "L" && element.getPageElementType() === SlidesApp.PageElementType.LINE) element.asLine().getLineFill().setSolidFill(theme);
    else if (target === "L" && element.getPageElementType() === SlidesApp.PageElementType.SHAPE) element.asShape().getBorder().getLineFill().setSolidFill(theme);
    else if (target === "F" && element.getPageElementType() === SlidesApp.PageElementType.SHAPE) element.asShape().getFill().setSolidFill(theme);
  });
  return { message: `Applied theme ${themeName}.` };
}

function solidHex(color: GoogleAppsScript.Slides.Color, scheme: GoogleAppsScript.Slides.ColorScheme): string {
  if (color.getColorType() === SlidesApp.ColorType.THEME) return scheme.getConcreteColor(color.asThemeColor().getThemeColorType()).asRgbColor().asHexString();
  return color.asRgbColor().asHexString();
}

function nearestTheme(hex: string, swatches: ThemeSwatch[]): string {
  const rgb = [1, 3, 5].map((start) => Number.parseInt(hex.slice(start, start + 2), 16));
  let best = swatches[0]!;
  let distance = Number.POSITIVE_INFINITY;
  swatches.forEach((swatch) => {
    const candidate = [1, 3, 5].map((start) => Number.parseInt(swatch.hex.slice(start, start + 2), 16));
    const next = rgb.reduce((sum, channel, index) => sum + (channel! - candidate[index]!) ** 2, 0);
    if (next < distance) { distance = next; best = swatch; }
  });
  return best.name;
}

export function convertSelectionColors(toTheme: boolean): { message: string } {
  const context = activeContext(1);
  const scheme = context.slide.getColorScheme();
  const swatches = currentThemeSwatches();
  const mapping = themeMap();
  let count = 0;
  visit(context.elements, (element) => {
    const convertFill = (
      solid: GoogleAppsScript.Slides.SolidFill,
      setter: (value: string | ThemeColorType) => void,
    ): void => {
      const color = solid.getColor();
      if (toTheme && color.getColorType() === SlidesApp.ColorType.RGB) setter(mapping[nearestTheme(color.asRgbColor().asHexString(), swatches)]!);
      else if (!toTheme && color.getColorType() === SlidesApp.ColorType.THEME) setter(solidHex(color, scheme));
      else return;
      count += 1;
    };
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      const fill = shape.getFill().getSolidFill();
      if (fill) convertFill(fill, (value) => shape.getFill().setSolidFill(value as ThemeColorType));
      const lineFill = shape.getBorder().getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) convertFill(lineFill.getSolidFill(), (value) => lineFill.setSolidFill(value as ThemeColorType));
      const textStyle = shape.getText().getTextStyle();
      const textColor = textStyle.getForegroundColor();
      if (textColor) {
        const pseudo = { getColor: () => textColor } as GoogleAppsScript.Slides.SolidFill;
        convertFill(pseudo, (value) => textStyle.setForegroundColor(value as ThemeColorType));
      }
    } else if (element.getPageElementType() === SlidesApp.PageElementType.LINE) {
      const lineFill = element.asLine().getLineFill();
      if (lineFill.getFillType() === SlidesApp.LineFillType.SOLID) convertFill(lineFill.getSolidFill(), (value) => lineFill.setSolidFill(value as ThemeColorType));
    }
  });
  return { message: `Converted ${count} color${count === 1 ? "" : "s"} to ${toTheme ? "theme links" : "RGB"}.` };
}

export function applyPaletteToTheme(paletteName: string): { message: string } {
  const palette = PALETTES[paletteName];
  if (!palette) throw new Error(`Unknown palette: ${paletteName}`);
  const context = activeContext();
  const accents = [SlidesApp.ThemeColorType.ACCENT1, SlidesApp.ThemeColorType.ACCENT2, SlidesApp.ThemeColorType.ACCENT3, SlidesApp.ThemeColorType.ACCENT4, SlidesApp.ThemeColorType.ACCENT5, SlidesApp.ThemeColorType.ACCENT6];
  context.presentation.getSlides().forEach((slide) => {
    const scheme = slide.getColorScheme();
    accents.forEach((accent, index) => scheme.setConcreteColor(accent, palette[index]!));
  });
  return { message: `Applied ${paletteName} to the deck's six theme accents.` };
}
