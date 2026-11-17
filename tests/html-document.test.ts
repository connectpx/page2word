import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import {
  decodeHtmlFile,
  normalizeHeadingLevels,
  normalizePastedText,
  normalizeTextNodes,
  sectionizeArticleHtml,
} from "../src/shared/html-document";
import { sanitizeHtml } from "../src/shared/sanitize";

describe("HTML document import", () => {
  it("normalizes pasted punctuation and removes invalid invisible characters", () => {
    expect(normalizePastedText("K\u201312 — “quote”…\uFFFD\u200B")).toBe(
      'K\u201312 - "quote"...',
    );
  });

  it("normalizes text nodes in imported HTML", () => {
    const dom = new JSDOM("<p>K\u201312 — “quote”\uFFFD</p>");

    normalizeTextNodes(dom.window.document.body);

    expect(dom.window.document.body.textContent).toBe('K\u201312 - "quote"');
    dom.window.close();
  });

  it("decodes a Windows-1252 HTML file using its declared charset", async () => {
    const source =
      '<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">' +
      "<p>K-12 Education</p>";
    const bytes = new TextEncoder().encode(source);
    bytes[source.indexOf("K-12") + 1] = 0x96;

    const html = await decodeHtmlFile(new Blob([bytes]));

    expect(html).toContain("K\u201312 Education");
    expect(html).not.toContain("\uFFFD");
  });

  it("splits Word-style H1 sections and keeps H2 subheadings with their parent", () => {
    const source = new JSDOM(
      "<div class=WordSection1><p>Document title</p>" +
        "<h1>Abstract</h1><p>Summary text.</p>" +
        "<h1>1. Introduction</h1><p>Introduction text.</p>" +
        "<h2>1.1. Background</h2><p>Background text.</p>" +
        "<h1>2. Theoretical and\nPractical Methods</h1><p>Methods text.</p></div>",
    );
    const html = sanitizeHtml(
      source.window.document.body.innerHTML,
      source.window.document,
    );
    const parsed = new JSDOM("");
    const document = parsed.window.document;
    document.body.innerHTML = html;
    normalizeHeadingLevels(document.body, document);
    const sections = sectionizeArticleHtml(
      document.body.innerHTML,
      document,
      "Imported paper",
    );

    expect(sections.map((section) => section.title)).toEqual([
      "Abstract",
      "1. Introduction",
      "2. Theoretical and Practical Methods",
    ]);
    expect(sections[1].html).toContain("<h3>1.1. Background</h3>");
    expect(sections[1].html).toContain("Background text.");
    source.window.close();
    parsed.window.close();
  });
});
