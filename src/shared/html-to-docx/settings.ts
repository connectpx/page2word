import type { DocxSettings } from "./types";

export const DOCX_SETTINGS_KEY = "docxSettings";

export const DEFAULT_DOCX_SETTINGS: DocxSettings = {
  fontFamily: "Times New Roman",
  fontSizePt: 12,
  titleFontSizePt: 16,
  headingFontSizePt: 14,
  subheadingFontSizePt: 13,
  tableFontSizePt: 10,
  lineSpacing: 1.5,
  spaceAfterPt: 8,
  alignment: "justify",
  pageSize: "Letter",
  marginsMm: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  includeTitle: true,
  includeSource: false,
  keepOriginalLinks: false,
  includeToc: false,
  tocLevels: 3,
  pageNumbers: false,
  pageBreakBeforeH1: false,
  keepHeadingsWithNext: true,
};

export function resolveDocxSettings(value: unknown): DocxSettings {
  if (!isRecord(value)) {
    return DEFAULT_DOCX_SETTINGS;
  }
  const candidate = isRecord(value.export) ? value.export : value;
  const margins = isRecord(candidate.marginsMm) ? candidate.marginsMm : {};
  const defaultMargins = DEFAULT_DOCX_SETTINGS.marginsMm;

  return {
    fontFamily: boundedText(
      candidate.fontFamily,
      DEFAULT_DOCX_SETTINGS.fontFamily,
      80,
    ),
    fontSizePt: boundedNumber(
      candidate.fontSizePt,
      DEFAULT_DOCX_SETTINGS.fontSizePt,
      6,
      72,
    ),
    titleFontSizePt: boundedNumber(
      candidate.titleFontSizePt,
      DEFAULT_DOCX_SETTINGS.titleFontSizePt,
      6,
      72,
    ),
    headingFontSizePt: boundedNumber(
      candidate.headingFontSizePt,
      DEFAULT_DOCX_SETTINGS.headingFontSizePt,
      6,
      72,
    ),
    subheadingFontSizePt: boundedNumber(
      candidate.subheadingFontSizePt,
      DEFAULT_DOCX_SETTINGS.subheadingFontSizePt,
      6,
      72,
    ),
    tableFontSizePt: boundedNumber(
      candidate.tableFontSizePt,
      DEFAULT_DOCX_SETTINGS.tableFontSizePt,
      6,
      72,
    ),
    lineSpacing: boundedNumber(
      candidate.lineSpacing,
      DEFAULT_DOCX_SETTINGS.lineSpacing,
      0.8,
      3,
    ),
    spaceAfterPt: boundedNumber(
      candidate.spaceAfterPt,
      DEFAULT_DOCX_SETTINGS.spaceAfterPt,
      0,
      72,
    ),
    alignment:
      candidate.alignment === "justify" || candidate.alignment === "left"
        ? candidate.alignment
        : DEFAULT_DOCX_SETTINGS.alignment,
    pageSize: candidate.pageSize === "A4" ? "A4" : "Letter",
    marginsMm: {
      top: boundedNumber(margins.top, defaultMargins.top, 0, 60),
      right: boundedNumber(margins.right, defaultMargins.right, 0, 60),
      bottom: boundedNumber(margins.bottom, defaultMargins.bottom, 0, 60),
      left: boundedNumber(margins.left, defaultMargins.left, 0, 60),
    },
    includeTitle: booleanValue(candidate.includeTitle, true),
    includeSource: booleanValue(candidate.includeSource, true),
    keepOriginalLinks: booleanValue(candidate.keepOriginalLinks, false),
    includeToc: booleanValue(candidate.includeToc, false),
    tocLevels: validTocLevels(candidate.tocLevels),
    pageNumbers: booleanValue(candidate.pageNumbers, true),
    pageBreakBeforeH1: booleanValue(candidate.pageBreakBeforeH1, false),
    keepHeadingsWithNext: booleanValue(candidate.keepHeadingsWithNext, true),
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function boundedText(
  value: unknown,
  fallback: string,
  maxLength: number,
): string {
  return typeof value === "string" && value.trim()
    ? value.trim().slice(0, maxLength)
    : fallback;
}

function boundedNumber(
  value: unknown,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  return typeof value === "number" &&
    Number.isFinite(value) &&
    value >= minimum &&
    value <= maximum
    ? value
    : fallback;
}

function booleanValue(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function validTocLevels(value: unknown): 1 | 2 | 3 | 4 {
  return value === 1 || value === 2 || value === 3 || value === 4 ? value : 3;
}
