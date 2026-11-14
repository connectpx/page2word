export interface ExtractedDoc {
  id: string;
  title: string;
  byline?: string;
  siteName?: string;
  sourceUrl: string;
  extractedAt: string;
  lang?: string;
  html: string;
  sections: { id: string; title: string; words: number }[];
  contentSections: ContentSection[];
  assets: ExtractedAsset[];
  stats: { words: number; images: number; tables: number; charts: number };
  warnings: string[];
}

export interface ExtractedAsset {
  id: string;
  mime: "image/png" | "image/jpeg" | "image/gif";
  width: number;
  height: number;
  bytes: number;
}

export interface ContentSection {
  id: string;
  title: string;
  words: number;
  html: string;
}

export interface StoredDocument {
  id: string;
  title: string;
  byline?: string;
  siteName?: string;
  sourceUrl: string;
  extractedAt: string;
  lang?: string;
  sectionCount: number;
  sectionsAvailable?: boolean;
  sections: Omit<ContentSection, "html">[];
  assets: ExtractedDoc["assets"];
  stats: ExtractedDoc["stats"];
  warnings: string[];
  settings: Record<string, unknown>;
}
