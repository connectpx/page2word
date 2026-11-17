import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import { sanitizeHtml } from "../src/shared/sanitize";

describe("HTML sanitization", () => {
  it("removes scripts and extension UI while preserving marked text", () => {
    const dom = new JSDOM("");
    const clean = sanitizeHtml(
      '<p onclick="alert(1)">Keep <span class="gr_1" data-gr-ext-installed>this</span> text.</p>' +
        "<grammarly-extension>toolbar</grammarly-extension><script>bad()</script>",
      dom.window.document,
    );

    expect(clean).toContain("Keep this text.");
    expect(clean).not.toMatch(/onclick|script|toolbar/);
    dom.window.close();
  });

  it("rejects unsafe links and image sources", () => {
    const dom = new JSDOM("");
    const clean = sanitizeHtml(
      '<p><a href="javascript:alert(1)">link</a><img src="https://example.test/photo.png" onerror="bad()"></p>',
      dom.window.document,
    );

    expect(clean).toContain("<a>link</a>");
    expect(clean).toContain('src="https://example.test/photo.png"');
    expect(clean).not.toContain("onerror");
    dom.window.close();
  });

  it("preserves table spans and captured-image references", () => {
    const dom = new JSDOM("");
    const clean = sanitizeHtml(
      '<table><tr><th rowspan="2" colspan="3">Header</th></tr></table>' +
        '<img src="p2w-asset:asset-one" data-p2w-asset="asset-one" width="100" height="50">',
      dom.window.document,
    );

    expect(clean).toContain('rowspan="2"');
    expect(clean).toContain('colspan="3"');
    expect(clean).toContain('data-p2w-asset="asset-one"');
    expect(clean).toContain('src="p2w-asset:asset-one"');
    expect(clean).toContain('width="100"');
    expect(clean).toContain('height="50"');
    dom.window.close();
  });

  it("preserves supported text styles and removes unsafe style declarations", () => {
    const dom = new JSDOM("");
    const clean = sanitizeHtml(
      '<p><span style="font-family: Arial; font-size: 18pt; font-weight: 700; color: #123456; background-image: url(javascript:bad())">Styled</span></p>',
      dom.window.document,
    );

    expect(clean).toContain("font-family: Arial");
    expect(clean).toContain("font-size: 18pt");
    expect(clean).toContain("font-weight: 700");
    expect(clean).toContain("color: rgb(18, 52, 86)");
    expect(clean).not.toContain("background-image");
    expect(clean).not.toContain("javascript:");
    dom.window.close();
  });
});
