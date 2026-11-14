import {
  AlignmentType,
  Bookmark,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  HeadingLevel,
  ImageRun,
  InternalHyperlink,
  LevelFormat,
  PageNumber,
  Paragraph,
  Packer,
  TableOfContents,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  type IParagraphOptions,
} from "docx";
import type {
  DocxBlock,
  DocxImageAsset,
  DocxMetadata,
  DocxParagraphBlock,
  DocxSettings,
  InlineRun,
  DocxTableBlock,
} from "./types";

const TWIPS_PER_MM = 1440 / 25.4;
const TWIPS_PER_POINT = 20;

export function createDocxDocument(
  blocks: DocxBlock[],
  metadata: DocxMetadata,
  settings: DocxSettings,
  assets: ReadonlyMap<string, DocxImageAsset> = new Map(),
): Document {
  const children: (Paragraph | Table | TableOfContents)[] = [];
  if (settings.includeTitle && metadata.title.trim()) {
    children.push(
      new Paragraph({
        text: metadata.title.trim(),
        heading: HeadingLevel.TITLE,
      }),
    );
  }
  if (settings.includeSource) {
    const sourceLines = [
      metadata.byline ? `By ${metadata.byline}` : "",
      metadata.extractedAt
        ? new Date(metadata.extractedAt).toLocaleDateString()
        : "",
      metadata.sourceUrl && settings.keepOriginalLinks
        ? `Source: ${metadata.sourceUrl}`
        : "",
    ].filter(Boolean);
    if (sourceLines.length) {
      children.push(
        new Paragraph({
          text: sourceLines.join("  ·  "),
          style: "SourceBlock",
        }),
      );
    }
  }
  if (settings.includeToc) {
    children.push(new Paragraph({ text: "Contents", style: "Title" }));
    children.push(
      new TableOfContents("Contents", {
        hyperlink: true,
        headingStyleRange: `1-${settings.tocLevels}`,
      }),
    );
  }

  for (const block of blocks) {
    if (block.kind === "toc") {
      children.push(
        new TableOfContents("Contents", {
          hyperlink: true,
          headingStyleRange: `1-${settings.tocLevels}`,
        }),
      );
      continue;
    }
    if (block.kind === "table") {
      if (block.bookmark) {
        children.push(
          new Paragraph({
            children: [new Bookmark({ id: block.bookmark, children: [] })],
            spacing: { before: 0, after: 0 },
          }),
        );
      }
      children.push(tableFromBlock(block, settings, assets));
    } else {
      children.push(paragraphFromBlock(block, settings, assets));
    }
  }

  const size =
    settings.pageSize === "A4"
      ? { width: 11906, height: 16838 }
      : { width: 12240, height: 15840 };
  const footer = settings.pageNumbers
    ? new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun("Page "),
              new TextRun({ children: [PageNumber.CURRENT] }),
            ],
          }),
        ],
      })
    : undefined;

  return new Document({
    title: metadata.title,
    subject: settings.keepOriginalLinks ? metadata.sourceUrl : undefined,
    creator: "Page2Word",
    features: { updateFields: settings.includeToc },
    numbering: {
      config: [
        {
          reference: "p2w-bullets",
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: LevelFormat.BULLET,
            text: ["•", "◦", "▪"][level % 3],
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: { left: 720 + level * 360, hanging: 360 },
              },
            },
          })),
        },
        {
          reference: "p2w-numbers",
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: [
              LevelFormat.DECIMAL,
              LevelFormat.LOWER_LETTER,
              LevelFormat.LOWER_ROMAN,
            ][level % 3],
            text: `%${level + 1}.`,
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: { left: 720 + level * 360, hanging: 360 },
              },
            },
          })),
        },
      ],
    },
    styles: {
      default: {
        document: {
          run: {
            font: settings.fontFamily,
            size: Math.round(settings.fontSizePt * 2),
          },
          paragraph: {
            spacing: {
              line: Math.round(settings.lineSpacing * 240),
              after: Math.round(settings.spaceAfterPt * TWIPS_PER_POINT),
            },
            alignment:
              settings.alignment === "justify"
                ? AlignmentType.BOTH
                : AlignmentType.LEFT,
          },
        },
        title: {
          run: {
            font: settings.fontFamily,
            size: Math.round(settings.titleFontSizePt * 2),
            bold: true,
            color: "202124",
          },
          paragraph: {
            alignment: AlignmentType.CENTER,
            spacing: { after: 240, line: 360 },
          },
        },
        heading1: headingStyle(
          settings.fontFamily,
          Math.round(settings.headingFontSizePt * 2),
          "202124",
        ),
        heading2: headingStyle(
          settings.fontFamily,
          Math.round(settings.subheadingFontSizePt * 2),
          "202124",
        ),
        heading3: headingStyle(
          settings.fontFamily,
          Math.round(settings.subheadingFontSizePt * 2),
          "202124",
        ),
        heading4: headingStyle(
          settings.fontFamily,
          Math.round(settings.subheadingFontSizePt * 2),
          "202124",
        ),
      },
      paragraphStyles: [
        {
          id: "SourceBlock",
          name: "Source Block",
          run: { font: settings.fontFamily, size: 18, color: "5F6368" },
          paragraph: { spacing: { after: 320 } },
        },
        {
          id: "Quote",
          name: "Quote",
          run: { font: settings.fontFamily, italics: true, color: "4B5563" },
          paragraph: {
            indent: { left: 480 },
            border: {
              left: { style: "single", size: 12, color: "9AA0A6", space: 8 },
            },
            spacing: { before: 120, after: 180 },
          },
        },
        {
          id: "Code",
          name: "Code",
          run: { font: "Courier New", size: 18 },
          paragraph: {
            indent: { left: 240 },
            spacing: { before: 40, after: 40, line: 240 },
          },
        },
        {
          id: "TableText",
          name: "Table Text",
          run: {
            font: settings.fontFamily,
            size: Math.round(settings.tableFontSizePt * 2),
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: { before: 0, after: 40, line: 360 },
          },
        },
        {
          id: "Caption",
          name: "Caption",
          run: {
            font: settings.fontFamily,
            size: Math.round(settings.fontSizePt * 2),
            italics: true,
            color: "5F6368",
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: { before: 80, after: 160 },
          },
        },
        {
          id: "FigureCaption",
          name: "Figure Caption",
          run: {
            font: settings.fontFamily,
            size: Math.round(settings.fontSizePt * 2),
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: { before: 80, after: 160, line: 360 },
            border: {
              bottom: {
                style: BorderStyle.SINGLE,
                size: 6,
                color: "202124",
              },
            },
          },
        },
        {
          id: "SourceBlockParagraph",
          name: "Source Block Paragraph",
          basedOn: "SourceBlock",
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size,
            margin: {
              top: Math.round(settings.marginsMm.top * TWIPS_PER_MM),
              right: Math.round(settings.marginsMm.right * TWIPS_PER_MM),
              bottom: Math.round(settings.marginsMm.bottom * TWIPS_PER_MM),
              left: Math.round(settings.marginsMm.left * TWIPS_PER_MM),
              header: 720,
              footer: 720,
            },
          },
        },
        ...(footer ? { footers: { default: footer } } : {}),
        children,
      },
    ],
  });
}

export async function packDocx(document: Document): Promise<Blob> {
  return Packer.toBlob(document);
}

function paragraphFromBlock(
  block: DocxParagraphBlock,
  settings: DocxSettings,
  assets: ReadonlyMap<string, DocxImageAsset>,
): Paragraph {
  const heading = headingForStyle(block.style);
  const inline = inlineChildren(block.runs, assets);
  const children = block.bookmark
    ? [new Bookmark({ id: block.bookmark, children: inline })]
    : inline;
  const options: IParagraphOptions = {
    children,
    ...(heading ? { heading } : { style: styleId(block.style) }),
    ...(block.list
      ? {
          numbering: {
            reference:
              block.list.type === "bullet" ? "p2w-bullets" : "p2w-numbers",
            level: block.list.level,
            instance: block.list.instance,
          },
        }
      : {}),
    ...(heading && settings.keepHeadingsWithNext ? { keepNext: true } : {}),
    ...(block.style === "heading1" && settings.pageBreakBeforeH1
      ? { pageBreakBefore: true }
      : {}),
    ...(block.style === "body" && settings.alignment === "justify"
      ? { alignment: AlignmentType.BOTH }
      : {}),
    ...(block.style === "body" && settings.alignment === "left"
      ? { alignment: AlignmentType.LEFT }
      : {}),
    ...(block.style === "rule" ? { thematicBreak: true } : {}),
    ...(block.style === "image"
      ? {
          alignment: AlignmentType.CENTER,
          spacing: { before: 240, after: 120 },
        }
      : {}),
    ...(block.style === "caption" ? { alignment: AlignmentType.LEFT } : {}),
  };
  return new Paragraph(options);
}

function inlineChildren(
  runs: InlineRun[],
  assets: ReadonlyMap<string, DocxImageAsset>,
): (TextRun | ExternalHyperlink | InternalHyperlink | ImageRun)[] {
  const result: (TextRun | ExternalHyperlink | InternalHyperlink | ImageRun)[] =
    [];
  for (let index = 0; index < runs.length;) {
    const run = runs[index];
    if (run.assetId) {
      const asset = assets.get(run.assetId);
      if (asset) {
        result.push(createImageRun(run, asset));
      } else {
        result.push(
          new TextRun({
            text: `[Image unavailable: ${run.alt || "image"}]`,
            italics: true,
          }),
        );
      }
      index += 1;
    } else if (run.anchor) {
      const anchor = run.anchor;
      const linked: TextRun[] = [];
      while (index < runs.length && runs[index].anchor === anchor) {
        linked.push(textRun(runs[index], true));
        index += 1;
      }
      result.push(new InternalHyperlink({ anchor, children: linked }));
    } else if (run.href) {
      const href = run.href;
      const linked: TextRun[] = [];
      while (index < runs.length && runs[index].href === href) {
        linked.push(textRun(runs[index], true));
        index += 1;
      }
      result.push(new ExternalHyperlink({ link: href, children: linked }));
    } else {
      result.push(textRun(run));
      index += 1;
    }
  }
  return result;
}

function tableFromBlock(
  block: DocxTableBlock,
  settings: DocxSettings,
  assets: ReadonlyMap<string, DocxImageAsset>,
): Table {
  const pageWidth = settings.pageSize === "A4" ? 11906 : 12240;
  const totalWidth = Math.max(
    1,
    pageWidth -
      Math.round(
        (settings.marginsMm.left + settings.marginsMm.right) * TWIPS_PER_MM,
      ),
  );
  const columns = Math.max(1, block.columnCount);
  const columnWidths = Array.from({ length: columns }, (_, index) =>
    index === columns - 1
      ? totalWidth - Math.floor(totalWidth / columns) * (columns - 1)
      : Math.floor(totalWidth / columns),
  );
  const border = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: "202124",
  };
  const noBorder = {
    style: BorderStyle.NONE,
    size: 0,
    color: "FFFFFF",
  };
  const rows = block.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const runs = cell.header
        ? cell.runs.map((run) => ({ ...run, bold: true }))
        : cell.runs;
      const children = inlineChildren(runs, assets);
      const options = {
        children: [
          new Paragraph({
            children: children.length ? children : [new TextRun("")],
            style: "TableText",
          }),
        ],
        width: {
          size: columnWidths
            .slice(0, cell.columnSpan)
            .reduce((sum, width) => sum + width, 0),
          type: WidthType.DXA,
        },
        ...(cell.columnSpan > 1 ? { columnSpan: cell.columnSpan } : {}),
        ...(cell.rowSpan > 1 ? { rowSpan: cell.rowSpan } : {}),
        borders: {
          top: row.header ? border : noBorder,
          bottom: row.header ? border : noBorder,
          left: noBorder,
          right: noBorder,
        },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
      };
      return new TableCell(options);
    });
    return new TableRow({
      children: cells.length
        ? cells
        : [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun("")] })],
            }),
          ],
      tableHeader: row.header,
      cantSplit: true,
    });
  });
  return new Table({
    rows,
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths,
    layout: "autofit",
    borders: {
      top: noBorder,
      bottom: noBorder,
      left: noBorder,
      right: noBorder,
      insideHorizontal: noBorder,
      insideVertical: noBorder,
    },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
  });
}

function createImageRun(run: InlineRun, asset: DocxImageAsset): ImageRun {
  const originalWidth = asset.width || run.width || 640;
  const originalHeight = asset.height || run.height || 480;
  const scale = Math.min(1, 620 / originalWidth, 800 / originalHeight);
  const width = Math.max(1, Math.round(originalWidth * scale));
  const height = Math.max(1, Math.round(originalHeight * scale));
  return new ImageRun({
    type:
      asset.mime === "image/jpeg"
        ? "jpg"
        : asset.mime === "image/gif"
          ? "gif"
          : "png",
    data: asset.bytes,
    transformation: { width, height },
    altText: {
      name: run.alt || "Article image",
      description: run.alt || "Article image",
      title: run.alt || "Article image",
    },
  });
}

function textRun(run: InlineRun, hyperlink = false): TextRun {
  return new TextRun({
    text: run.lineBreak ? undefined : run.text,
    break: run.lineBreak ? 1 : undefined,
    bold: run.bold,
    italics: run.italic,
    underline: run.underline ? {} : undefined,
    strike: run.strike,
    subScript: run.subscript,
    superScript: run.superscript,
    font: run.code ? "Courier New" : undefined,
    color: hyperlink ? "0563C1" : undefined,
  });
}

function headingStyle(
  font: string,
  size: number,
  color: string,
): {
  run: { font: string; size: number; bold: true; color: string };
  paragraph: { spacing: { before: number; after: number } };
} {
  return {
    run: { font, size, bold: true, color },
    paragraph: { spacing: { before: 240, after: 100 } },
  };
}

function headingForStyle(
  style: DocxParagraphBlock["style"],
): (typeof HeadingLevel)[keyof typeof HeadingLevel] | undefined {
  switch (style) {
    case "heading1":
      return HeadingLevel.HEADING_1;
    case "heading2":
      return HeadingLevel.HEADING_2;
    case "heading3":
      return HeadingLevel.HEADING_3;
    case "heading4":
      return HeadingLevel.HEADING_4;
    default:
      return undefined;
  }
}

function styleId(style: DocxParagraphBlock["style"]): string {
  switch (style) {
    case "quote":
      return "Quote";
    case "code":
      return "Code";
    case "rule":
      return "Normal";
    case "image":
      return "Normal";
    case "caption":
      return "Caption";
    case "figureCaption":
      return "FigureCaption";
    case "table":
      return "TableText";
    case "source":
      return "SourceBlock";
    default:
      return "Normal";
  }
}
