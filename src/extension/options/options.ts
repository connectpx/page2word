import {
  DEFAULT_DOCX_SETTINGS,
  DOCX_SETTINGS_KEY,
  resolveDocxSettings,
} from "../../shared/html-to-docx/settings";

const form = requiredField<HTMLFormElement>("#settings-form");
const statusElement = requiredField<HTMLElement>("#settings-status");
const resetButton = requiredField<HTMLButtonElement>("#reset-settings");

const fields = {
  fontFamily: requiredField<HTMLSelectElement>("#font-family"),
  fontSizePt: requiredField<HTMLInputElement>("#font-size-pt"),
  titleFontSizePt: requiredField<HTMLInputElement>("#title-font-size-pt"),
  headingFontSizePt: requiredField<HTMLInputElement>("#heading-font-size-pt"),
  subheadingFontSizePt: requiredField<HTMLInputElement>(
    "#subheading-font-size-pt",
  ),
  tableFontSizePt: requiredField<HTMLInputElement>("#table-font-size-pt"),
  lineSpacing: requiredField<HTMLInputElement>("#line-spacing"),
  alignment: requiredField<HTMLSelectElement>("#alignment"),
  originalLinks: requiredField<HTMLSelectElement>("#original-links"),
};

form.addEventListener("submit", (event) => {
  event.preventDefault();
  void saveSettings(readFormSettings()).catch(showSettingsError);
});

resetButton.addEventListener("click", () => {
  fillForm(DEFAULT_DOCX_SETTINGS);
  void saveSettings(DEFAULT_DOCX_SETTINGS).catch(showSettingsError);
});

void loadSettings().catch(showSettingsError);

async function loadSettings(): Promise<void> {
  const stored = await chrome.storage.local.get(DOCX_SETTINGS_KEY);
  fillForm(resolveDocxSettings(stored[DOCX_SETTINGS_KEY]));
}

async function saveSettings(
  settings: ReturnType<typeof resolveDocxSettings>,
): Promise<void> {
  statusElement.textContent = "Saving…";
  await chrome.storage.local.set({ [DOCX_SETTINGS_KEY]: settings });
  statusElement.textContent =
    "Settings saved. They apply to the current editor and future extractions.";
}

function readFormSettings(): ReturnType<typeof resolveDocxSettings> {
  return resolveDocxSettings({
    ...DEFAULT_DOCX_SETTINGS,
    fontFamily: fields.fontFamily.value,
    fontSizePt: Number(fields.fontSizePt.value),
    titleFontSizePt: Number(fields.titleFontSizePt.value),
    headingFontSizePt: Number(fields.headingFontSizePt.value),
    subheadingFontSizePt: Number(fields.subheadingFontSizePt.value),
    tableFontSizePt: Number(fields.tableFontSizePt.value),
    lineSpacing: Number(fields.lineSpacing.value),
    alignment: fields.alignment.value,
    keepOriginalLinks: fields.originalLinks.value === "keep",
  });
}

function fillForm(settings: ReturnType<typeof resolveDocxSettings>): void {
  fields.fontFamily.value = settings.fontFamily;
  fields.fontSizePt.value = String(settings.fontSizePt);
  fields.titleFontSizePt.value = String(settings.titleFontSizePt);
  fields.headingFontSizePt.value = String(settings.headingFontSizePt);
  fields.subheadingFontSizePt.value = String(settings.subheadingFontSizePt);
  fields.tableFontSizePt.value = String(settings.tableFontSizePt);
  fields.lineSpacing.value = String(settings.lineSpacing);
  fields.alignment.value = settings.alignment;
  fields.originalLinks.value = settings.keepOriginalLinks ? "keep" : "remove";
}

function requiredField<T extends HTMLElement>(selector: string): T {
  const field = document.querySelector<T>(selector);
  if (!field) {
    throw new Error(`Missing settings field: ${selector}`);
  }
  return field;
}

function showSettingsError(error: unknown): void {
  if (statusElement) {
    statusElement.textContent =
      error instanceof Error ? error.message : "Could not save settings.";
  }
}
