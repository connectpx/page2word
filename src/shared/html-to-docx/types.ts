export type ParagraphStyle =
  | "body"
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "quote"
  | "code"
  | "rule"
  | "image"
  | "caption"
  | "figureCaption"
  | "table"
  | "source";

export interface InlineRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  subscript?: boolean;
  superscript?: boolean;
  code?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  href?: string;
  anchor?: string;
  lineBreak?: boolean;
  assetId?: string;
  alt?: string;
  width?: number;
  height?: number;
}

export interface DocxTableCell {
  runs: InlineRun[];
  columnSpan: number;
  rowSpan: number;
  header: boolean;
  alignment?: "left" | "center" | "right";
  paddingPt?: number;
  border?: boolean;
}

export interface DocxTableRow {
  cells: DocxTableCell[];
  header: boolean;
}

export interface DocxTableBlock {
  kind: "table";
  rows: DocxTableRow[];
  columnCount: number;
  caption?: string;
  bookmark?: string;
}

export interface DocxParagraphBlock {
  kind: "paragraph";
  style: ParagraphStyle;
  runs: InlineRun[];
  bookmark?: string;
  list?: {
    type: "bullet" | "number";
    level: number;
    instance: number;
  };
  pageBreakBefore?: boolean;
}

export interface DocxTocBlock {
  kind: "toc";
}

export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxTocBlock;

export interface DocxImageAsset {
  id: string;
  mime: "image/png" | "image/jpeg" | "image/gif";
  width: number;
  height: number;
  bytes: ArrayBuffer;
}

export interface DocxMetadata {
  title: string;
  sourceUrl?: string;
  byline?: string;
  extractedAt?: string;
  lang?: string;
}

export interface DocxSettings {
  fontFamily: string;
  fontSizePt: number;
  titleFontSizePt: number;
  headingFontSizePt: number;
  subheadingFontSizePt: number;
  tableFontSizePt: number;
  tableLineSpacing: number;
  lineSpacing: number;
  spaceAfterPt: number;
  alignment: "left" | "justify";
  pageSize: "A4" | "Letter";
  marginsMm: { top: number; right: number; bottom: number; left: number };
  includeTitle: boolean;
  includeSource: boolean;
  keepOriginalLinks: boolean;
  includeToc: boolean;
  tocLevels: 1 | 2 | 3 | 4;
  pageNumbers: boolean;
  pageBreakBeforeH1: boolean;
  keepHeadingsWithNext: boolean;
  excludedSectionKeywords: string[];
}

export interface ExportProgress {
  stage: "preparing" | "building";
  completed: number;
  total: number;
}
