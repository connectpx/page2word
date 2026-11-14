import { inflateRawSync } from "node:zlib";
import { JSDOM } from "jsdom";
import { Document, Packer } from "docx";
import { describe, expect, it } from "vitest";
import { createDocxDocument } from "../src/shared/html-to-docx/document";
import { parseHtmlToDocxBlocks } from "../src/shared/html-to-docx/parse";
import {
  DEFAULT_DOCX_SETTINGS,
  resolveDocxSettings,
} from "../src/shared/html-to-docx/settings";

describe("HTML to DOCX export", () => {
  it("removes original links by default and creates internal figure/table links", async () => {
    const dom = new JSDOM("");
    const blocks = await parseHtmlToDocxBlocks(
      [
        '<p>See <a href="#p2w_figure_1">Fig. 1</a> and ' +
          '<a href="#p2w_table_1">Table 1</a>. <a href="https://example.test">source</a>.</p>' +
          '<figure id="p2w_figure_1"><figcaption>Fig. 1. Diagram</figcaption></figure>' +
          '<table id="p2w_table_1"><tbody><tr><td>Value</td></tr></tbody></table>',
      ],
      dom.window.document,
    );

    expect(resolveDocxSettings(undefined).keepOriginalLinks).toBe(false);
    expect(blocks[0]).toMatchObject({
      runs: [
        { text: "See " },
        { text: "Fig. 1", anchor: "p2w_figure_1" },
        { text: " and " },
        { text: "Table 1", anchor: "p2w_table_1" },
        { text: ". " },
        { text: "source" },
        { text: "." },
      ],
    });
    expect(blocks).toContainEqual(
      expect.objectContaining({
        kind: "paragraph",
        style: "figureCaption",
        bookmark: "p2w_figure_1",
      }),
    );
    expect(blocks).toContainEqual(
      expect.objectContaining({
        kind: "table",
        bookmark: "p2w_table_1",
      }),
    );

    const document = createDocxDocument(
      blocks,
      { title: "Internal references" },
      { ...DEFAULT_DOCX_SETTINGS, includeTitle: false },
    );
    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    expect(documentXml).toContain('w:anchor="p2w_figure_1"');
    expect(documentXml).toContain('w:anchor="p2w_table_1"');
    expect(documentXml).toContain('w:name="p2w_figure_1"');
    expect(documentXml).toContain('w:name="p2w_table_1"');
    expect(documentXml).not.toContain("https://example.test");
    dom.window.close();
  });

  it("preserves semantic text, formatting, links, nested lists, and sanitizes first", async () => {
    const dom = new JSDOM("");
    const blocks = await parseHtmlToDocxBlocks(
      [
        "<script>bad()</script><h2>Section heading</h2>" +
          '<p>A <strong>bold</strong> and <em>italic</em> <a href="https://example.test">link</a>.</p>' +
          "<ul><li>Outer<ul><li>Inner item</li></ul></li></ul>" +
          "<blockquote><p>Quoted words</p></blockquote>" +
          "<pre>first line\nsecond line</pre>",
      ],
      dom.window.document,
      undefined,
      undefined,
      { keepOriginalLinks: true },
    );

    expect(blocks.map((block) => block.kind)).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
      "paragraph",
      "paragraph",
      "paragraph",
      "paragraph",
    ]);
    expect(blocks[0]).toMatchObject({
      style: "heading2",
      runs: [{ text: "Section heading" }],
    });
    expect(blocks[1]).toMatchObject({
      runs: [
        { text: "A " },
        { text: "bold", bold: true },
        { text: " and " },
        { text: "italic", italic: true },
        { text: " " },
        { text: "link", href: "https://example.test" },
        { text: "." },
      ],
    });
    expect(blocks[2]).toMatchObject({
      list: { type: "bullet", level: 0 },
    });
    expect(blocks[3]).toMatchObject({
      list: { type: "bullet", level: 1 },
    });
    expect(blocks[4]).toMatchObject({ style: "quote" });
    expect(blocks[5]).toMatchObject({
      style: "code",
      runs: [{ text: "first line" }],
    });
    expect(blocks[6]).toMatchObject({
      style: "code",
      runs: [{ text: "second line" }],
    });
    expect(JSON.stringify(blocks)).not.toContain("bad()");
    dom.window.close();
  });

  it("creates a valid DOCX archive with heading styles, numbering, links, and page setup", async () => {
    const dom = new JSDOM("");
    const blocks = await parseHtmlToDocxBlocks(
      [
        '<h1>Introduction</h1><p>Read the <a href="https://example.test/article">source</a>.</p>' +
          "<ol><li>First</li><li>Second</li></ol>",
      ],
      dom.window.document,
      undefined,
      undefined,
      { keepOriginalLinks: true },
    );
    const document = createDocxDocument(
      blocks,
      {
        title: "Test article",
        sourceUrl: "https://example.test/article",
        byline: "Author",
        extractedAt: "2025-01-02T00:00:00.000Z",
      },
      { ...DEFAULT_DOCX_SETTINGS, includeToc: true },
    );
    expect(document).toBeInstanceOf(Document);

    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    const numberingXml = readZipEntry(archive, "word/numbering.xml");
    const stylesXml = readZipEntry(archive, "word/styles.xml");
    const relationshipsXml = readZipEntry(
      archive,
      "word/_rels/document.xml.rels",
    );
    const settingsXml = readZipEntry(archive, "word/settings.xml");

    expect(documentXml).toContain("Introduction");
    expect(documentXml).toContain('w:val="Heading1"');
    expect(documentXml).toContain("TOC");
    expect(documentXml).toContain("First");
    expect(documentXml).toContain("Second");
    expect(documentXml).toMatch(/<w:numPr>[\s\S]*?<w:ilvl w:val="0"\/>/);
    expect(numberingXml).toContain('<w:lvlText w:val="•"/>');
    expect(stylesXml).toContain("Times New Roman");
    expect(stylesXml).toContain("Heading1");
    expect(relationshipsXml).toContain("https://example.test/article");
    expect(settingsXml).toContain("updateFields");
    dom.window.close();
  });

  it("honors cancellation while preparing a large document", async () => {
    const dom = new JSDOM("");
    const controller = new AbortController();
    controller.abort();

    await expect(
      parseHtmlToDocxBlocks(
        ["<p>Cancelled</p>"],
        dom.window.document,
        undefined,
        controller.signal,
      ),
    ).rejects.toMatchObject({ name: "AbortError" });
    dom.window.close();
  });

  it("does not export a duplicate literal bullet from an unordered list", async () => {
    const dom = new JSDOM("");
    const blocks = await parseHtmlToDocxBlocks(
      ["<ul><li>• First item</li></ul>"],
      dom.window.document,
    );

    expect(blocks).toContainEqual({
      kind: "paragraph",
      style: "body",
      runs: [{ text: "First item" }],
      list: { type: "bullet", level: 0, instance: 1 },
    });
    const document = createDocxDocument(
      blocks,
      { title: "Bullet list" },
      DEFAULT_DOCX_SETTINGS,
    );
    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    const numberingXml = readZipEntry(archive, "word/numbering.xml");
    expect(documentXml).toContain("<w:numPr>");
    expect(documentXml).toContain("First item");
    expect(numberingXml).toContain('<w:lvlText w:val="•"/>');
    dom.window.close();
  });

  it("exports figure captions stored in nested paragraphs", async () => {
    const dom = new JSDOM("");
    const caption =
      "Fig. 2. Identity-citizenship-educational model: Main interrelationships expressed by the participants in the MCEfS program.";
    const blocks = await parseHtmlToDocxBlocks(
      [
        '<figure><span><img src="https://example.test/figure.jpg" alt="Identity diagram"></span>' +
          '<span class="captions text-s"><span><p><span class="label">Fig. 2</span>. ' +
          "Identity-citizenship-educational model: Main interrelationships expressed by the participants in the MCEfS program." +
          "</p></span></span></figure>",
      ],
      dom.window.document,
    );

    expect(blocks).toContainEqual({
      kind: "paragraph",
      style: "figureCaption",
      runs: [
        { text: "Fig.", bold: true },
        { text: caption.slice("Fig.".length) },
      ],
    });
    const document = createDocxDocument(
      blocks,
      { title: "Figure caption" },
      DEFAULT_DOCX_SETTINGS,
    );
    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    const stylesXml = readZipEntry(archive, "word/styles.xml");
    expect(documentXml).toContain('<w:t xml:space="preserve">Fig.</w:t>');
    expect(documentXml).toContain(
      '<w:t xml:space="preserve"> 2. Identity-citizenship-educational model: Main interrelationships expressed by the participants in the MCEfS program.</w:t>',
    );
    expect(documentXml).toContain("Fig.");
    expect(stylesXml).toMatch(
      /w:style[^>]*w:styleId="FigureCaption"[\s\S]*?w:pBdr[\s\S]*?w:bottom w:val="single"/,
    );
    expect(stylesXml).toMatch(
      /w:style[^>]*w:styleId="FigureCaption"[\s\S]*?w:sz w:val="24"/,
    );
    dom.window.close();
  });

  it("uses the requested Times New Roman formatting defaults", async () => {
    const settings = resolveDocxSettings(undefined);
    expect(settings).toMatchObject({
      fontFamily: "Times New Roman",
      fontSizePt: 12,
      titleFontSizePt: 16,
      headingFontSizePt: 14,
      subheadingFontSizePt: 13,
      tableFontSizePt: 10,
      lineSpacing: 1.5,
      alignment: "justify",
    });

    const document = createDocxDocument(
      [
        {
          kind: "paragraph",
          style: "heading1",
          runs: [{ text: "Heading" }],
        },
        {
          kind: "paragraph",
          style: "heading2",
          runs: [{ text: "Subheading" }],
        },
        { kind: "paragraph", style: "body", runs: [{ text: "Body" }] },
        {
          kind: "paragraph",
          style: "figureCaption",
          runs: [{ text: "Figure", bold: true }],
        },
      ],
      { title: "Main title" },
      settings,
    );
    const archive = await Packer.toBuffer(document);
    const stylesXml = readZipEntry(archive, "word/styles.xml");

    expect(stylesXml).toContain("Times New Roman");
    expect(stylesXml).toContain('w:sz w:val="24"');
    expect(stylesXml).toContain('w:sz w:val="32"');
    expect(stylesXml).toContain('w:sz w:val="28"');
    expect(stylesXml).toContain('w:sz w:val="26"');
    expect(stylesXml).toContain('w:sz w:val="20"');
    expect(stylesXml).toMatch(
      /w:style[^>]*w:styleId="Caption"[\s\S]*?w:sz w:val="24"/,
    );
    expect(stylesXml).toContain('w:line="360"');
    expect(stylesXml).toContain('w:val="both"');
    expect(stylesXml).toContain('w:val="center"');
  });

  it("applies validated font and page settings to the generated DOCX", async () => {
    const settings = resolveDocxSettings({
      fontFamily: "Times New Roman",
      fontSizePt: 12,
      pageSize: "A4",
      marginsMm: { top: 20, right: 18, bottom: 20, left: 18 },
    });
    const document = createDocxDocument(
      [
        {
          kind: "paragraph",
          style: "body",
          runs: [{ text: "Formatted text" }],
        },
      ],
      { title: "Formatted document" },
      settings,
    );
    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    const stylesXml = readZipEntry(archive, "word/styles.xml");

    expect(documentXml).toMatch(/w:pgSz[^>]*w:w="11906"[^>]*w:h="16838"/);
    expect(stylesXml).toContain("Times New Roman");
    expect(stylesXml).toContain('w:sz w:val="24"');
  });

  it("preserves table spans and embeds captured images in the DOCX package", async () => {
    const dom = new JSDOM("");
    const blocks = await parseHtmlToDocxBlocks(
      [
        '<table><thead><tr><th rowspan="2">Group</th><th colspan="2">Values</th></tr>' +
          "<tr><th>One</th><th>Two</th></tr></thead>" +
          "<tbody><tr><td>Observed</td><td>14</td><td>18</td></tr></tbody></table>" +
          '<p><img src="p2w-asset:asset-one" data-p2w-asset="asset-one" alt="Chart"></p>',
      ],
      dom.window.document,
    );
    const table = blocks.find((block) => block.kind === "table");
    expect(table?.kind).toBe("table");
    if (table?.kind !== "table") {
      throw new Error("The table was not parsed.");
    }
    expect(table.columnCount).toBe(3);
    expect(table.rows[0]).toMatchObject({
      header: true,
      cells: [
        { runs: [{ text: "Group" }], rowSpan: 2, header: true },
        { runs: [{ text: "Values" }], columnSpan: 2, header: true },
      ],
    });
    expect(table.rows[1]).toMatchObject({
      header: true,
      cells: [
        { runs: [{ text: "One" }], header: true },
        { runs: [{ text: "Two" }], header: true },
      ],
    });

    const png = new Uint8Array(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j6sgAAAAASUVORK5CYII=",
        "base64",
      ),
    );
    const assets = new Map([
      [
        "asset-one",
        {
          id: "asset-one",
          mime: "image/png" as const,
          width: 1,
          height: 1,
          bytes: png.buffer,
        },
      ],
    ]);
    const document = createDocxDocument(
      blocks,
      { title: "Tables and figures" },
      DEFAULT_DOCX_SETTINGS,
      assets,
    );
    const archive = await Packer.toBuffer(document);
    const documentXml = readZipEntry(archive, "word/document.xml");
    const stylesXml = readZipEntry(archive, "word/styles.xml");
    const relationshipsXml = readZipEntry(
      archive,
      "word/_rels/document.xml.rels",
    );

    expect(documentXml).toContain('<w:gridSpan w:val="2"/>');
    expect(documentXml).toContain('<w:vMerge w:val="restart"/>');
    expect(documentXml).toContain('<w:vMerge w:val="continue"/>');
    expect(documentXml).toContain("<w:tblHeader");
    expect(documentXml).not.toContain("<w:shd");
    expect(stylesXml).toMatch(
      /w:style[^>]*w:styleId="TableText"[\s\S]*?w:jc w:val="left"/,
    );
    expect(documentXml).toMatch(
      /<w:tcBorders>.*?<w:top w:val="single"[^>]*\/>.*?<w:left w:val="none"[^>]*\/>.*?<w:bottom w:val="single"[^>]*\/>/,
    );
    expect(documentXml).toContain('<w:insideH w:val="none"');
    expect(documentXml).toContain("<w:drawing>");
    expect(documentXml).toContain('w:before="240"');
    expect(relationshipsXml).toContain('Target="media/image1.png"');
    expect(readZipEntryBuffer(archive, "word/media/image1.png")).toEqual(
      Buffer.from(png),
    );
    dom.window.close();
  });
});

function readZipEntry(archive: Buffer, name: string): string {
  return readZipEntryBuffer(archive, name).toString("utf8");
}

function readZipEntryBuffer(archive: Buffer, name: string): Buffer {
  const endSignature = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
  const endOffset = archive.lastIndexOf(endSignature);
  if (endOffset < 0) {
    throw new Error("The generated DOCX is not a ZIP archive.");
  }
  const entryCount = archive.readUInt16LE(endOffset + 10);
  const directoryOffset = archive.readUInt32LE(endOffset + 16);
  let cursor = directoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (archive.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error("The DOCX ZIP directory is malformed.");
    }
    const method = archive.readUInt16LE(cursor + 10);
    const compressedSize = archive.readUInt32LE(cursor + 20);
    const nameLength = archive.readUInt16LE(cursor + 28);
    const extraLength = archive.readUInt16LE(cursor + 30);
    const commentLength = archive.readUInt16LE(cursor + 32);
    const localOffset = archive.readUInt32LE(cursor + 42);
    const entryName = archive
      .subarray(cursor + 46, cursor + 46 + nameLength)
      .toString("utf8");
    if (entryName === name) {
      const localNameLength = archive.readUInt16LE(localOffset + 26);
      const localExtraLength = archive.readUInt16LE(localOffset + 28);
      const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = archive.subarray(
        dataOffset,
        dataOffset + compressedSize,
      );
      const content =
        method === 8 ? inflateRawSync(compressed) : Buffer.from(compressed);
      return content;
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  throw new Error(`The DOCX archive is missing ${name}.`);
}
