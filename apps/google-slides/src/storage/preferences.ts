export interface ReferenceRecord {
  presentationId: string;
  slideId: string;
  objectId: string;
  label: string;
}

export interface UserSettings {
  palette: string;
  gapCm: number;
  matrixColumns: number;
  useAtomicUpdates: boolean;
}

const REFERENCE_KEY = "slideAid.reference";
const SETTINGS_KEY = "slideAid.settings";

export const PALETTES: Record<string, readonly string[]> = {
  Office: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
  "Nordic Blue": ["#1F4E79", "#2E75B6", "#9DC3E6", "#BDD7EE", "#636363", "#D9D9D9"],
  Fjord: ["#264653", "#2A9D8F", "#E9C46A", "#F4A261", "#E76F51", "#8AB17D"],
  Forest: ["#1B4332", "#2D6A4F", "#40916C", "#74C69D", "#B7E4C7", "#95D5B2"],
  Sunset: ["#073B4C", "#118AB2", "#06D6A0", "#FFD166", "#EF476F", "#26547C"],
  Berry: ["#4A1942", "#893168", "#C05299", "#E29ACD", "#6F6F6F", "#CFCFCF"],
  Greyscale: ["#212529", "#495057", "#6C757D", "#ADB5BD", "#CED4DA", "#DEE2E6"],
  Financial: ["#00304D", "#006BA6", "#FFB81C", "#97999B", "#DA291C", "#63666A"],
  Vivid: ["#3D348B", "#7678ED", "#F7B801", "#F18701", "#F35B04", "#5F0F40"],
};

const defaultSettings: UserSettings = { palette: "Office", gapCm: 0.2, matrixColumns: 3, useAtomicUpdates: true };

function userProperties(): GoogleAppsScript.Properties.Properties {
  return PropertiesService.getUserProperties();
}

export function getReference(): ReferenceRecord | null {
  const raw = userProperties().getProperty(REFERENCE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as ReferenceRecord;
  } catch {
    return null;
  }
}

export function saveReference(reference: ReferenceRecord): void {
  userProperties().setProperty(REFERENCE_KEY, JSON.stringify(reference));
}

export function deleteReference(): void {
  userProperties().deleteProperty(REFERENCE_KEY);
}

export function getSettings(): UserSettings {
  const raw = userProperties().getProperty(SETTINGS_KEY);
  if (!raw) return { ...defaultSettings };
  try {
    return { ...defaultSettings, ...(JSON.parse(raw) as Partial<UserSettings>) };
  } catch {
    return { ...defaultSettings };
  }
}

export function updateSettings(patch: Partial<UserSettings>): UserSettings {
  const settings = { ...getSettings(), ...patch };
  userProperties().setProperty(SETTINGS_KEY, JSON.stringify(settings));
  return settings;
}

export function currentPalette(): readonly string[] {
  const settings = getSettings();
  return PALETTES[settings.palette] ?? PALETTES.Office!;
}
