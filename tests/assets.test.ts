import { describe, expect, it, vi } from "vitest";
import { captureDocumentAssets } from "../src/extension/background/assets";
import type { ExtractedDoc } from "../src/shared/types";

describe("document image capture", () => {
  it("embeds supported image bytes and rewrites the article reference", async () => {
    const document = createDocument(
      '<figure><img src="https://example.test/chart.png" alt="Study chart" width="1" height="1"></figure>',
    );
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j6sgAAAAASUVORK5CYII=",
      "base64",
    );
    const fetchImage = vi.fn(
      async () =>
        new Response(png, {
          status: 200,
          headers: { "content-type": "image/png" },
        }),
    );

    const assets = await captureDocumentAssets(document, fetchImage);

    expect(assets).toHaveLength(1);
    expect(assets[0].metadata).toMatchObject({
      mime: "image/png",
      width: 1,
      height: 1,
      bytes: png.byteLength,
    });
    expect(Buffer.from(assets[0].base64, "base64")).toEqual(png);
    expect(document.contentSections[0].html).toContain(
      `src="p2w-asset:${assets[0].metadata.id}"`,
    );
    expect(document.contentSections[0].html).toContain(
      `data-p2w-asset="${assets[0].metadata.id}"`,
    );
    expect(document.warnings).toEqual([]);
  });

  it("labels unavailable images and records a warning", async () => {
    const document = createDocument(
      '<p><img src="https://example.test/missing.png" alt="Missing figure"></p>',
    );
    const fetchImage = vi.fn(async () => new Response(null, { status: 404 }));

    const assets = await captureDocumentAssets(document, fetchImage);

    expect(assets).toEqual([]);
    expect(document.contentSections[0].html).toContain(
      "[Image unavailable: Missing figure]",
    );
    expect(document.warnings).toHaveLength(1);
  });
});

function createDocument(html: string): ExtractedDoc {
  const contentSections = [
    { id: "section-1", title: "Article", words: 1, html },
  ];
  return {
    id: "document-test-123",
    title: "Article",
    sourceUrl: "https://example.test/article",
    extractedAt: "2026-11-12T00:00:00.000Z",
    html,
    sections: [{ id: "section-1", title: "Article", words: 1 }],
    contentSections,
    assets: [],
    stats: { words: 1, images: 1, tables: 0, charts: 0 },
    warnings: [],
  };
}
