import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  DOCX_SETTINGS_KEY,
  DEFAULT_DOCX_SETTINGS,
} from "../src/shared/html-to-docx/settings";
import {
  documentMetaKey,
  documentSectionKey,
} from "../src/shared/storage-keys";

const storageMocks = vi.hoisted(() => ({
  deleteEditorDocument: vi.fn(),
  loadEditorDocument: vi.fn(),
  saveEditorDocument: vi.fn(),
  saveEditorSections: vi.fn(),
  toEditorSection: vi.fn(
    (
      documentId: string,
      section: { id: string; title: string; words: number },
      html: string,
      done: boolean,
    ) => ({ documentId, ...section, html, done }),
  ),
}));

vi.mock("../src/editor/storage", () => storageMocks);

const documentId = "extension-editor-test";
const sectionId = "section-1";
const meta = {
  id: documentId,
  title: "Test article",
  sourceUrl: "https://example.test/article",
  extractedAt: "2026-11-12T00:00:00.000Z",
  sectionCount: 1,
  sectionsAvailable: true,
  sections: [{ id: sectionId, title: "Introduction", words: 2 }],
  assets: [],
  stats: { words: 2, images: 0, tables: 0, charts: 0 },
  warnings: [],
  settings: {},
};
const section = {
  id: sectionId,
  title: "Introduction",
  words: 2,
  html: "<p>Extension page content.</p>",
};

describe("bundled extension editor", () => {
  let localValues: Record<string, unknown>;
  let sendMessage: ReturnType<typeof vi.fn>;
  let dom: JSDOM;
  let storageChangedListener:
    | ((
        changes: Record<string, { newValue?: unknown }>,
        areaName: string,
      ) => void)
    | undefined;

  beforeEach(async () => {
    vi.resetModules();
    vi.unstubAllGlobals();
    vi.clearAllMocks();
    storageChangedListener = undefined;
    storageMocks.loadEditorDocument.mockResolvedValue({
      document: undefined,
      sections: [],
      assets: [],
    });
    storageMocks.saveEditorDocument.mockResolvedValue(undefined);
    storageMocks.saveEditorSections.mockResolvedValue(undefined);

    const editorHtml = readFileSync(
      resolve(
        dirname(fileURLToPath(import.meta.url)),
        "../src/editor/index.html",
      ),
      "utf8",
    );
    dom = new JSDOM(editorHtml, {
      url: `https://extension.test/editor/index.html#doc=${documentId}`,
    });
    localValues = {
      [documentMetaKey(documentId)]: meta,
      [documentSectionKey(documentId, 0)]: section,
    };
    sendMessage = vi.fn(async () => undefined);
    vi.stubGlobal("window", dom.window);
    vi.stubGlobal("document", dom.window.document);
    vi.stubGlobal("location", dom.window.location);
    vi.stubGlobal("CSS", { escape: (value: string) => value });
    vi.stubGlobal("chrome", {
      runtime: {
        id: "test-extension",
        sendMessage,
        openOptionsPage: vi.fn(async () => undefined),
      },
      storage: {
        onChanged: {
          addListener: vi.fn(
            (
              callback: (
                changes: Record<string, { newValue?: unknown }>,
                areaName: string,
              ) => void,
            ) => {
              storageChangedListener = callback;
            },
          ),
        },
        local: {
          get: vi.fn(async (keys: string | string[]) => {
            const requested = Array.isArray(keys) ? keys : [keys];
            return Object.fromEntries(
              requested
                .filter((key) => key in localValues)
                .map((key) => [key, localValues[key]]),
            );
          }),
        },
      },
    });
  });

  it("loads the stored article directly and acknowledges the handoff", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );

    expect(sendMessage).toHaveBeenCalledWith({
      type: "ACK",
      id: documentId,
    });
    expect(storageMocks.saveEditorDocument).toHaveBeenCalled();
    expect(
      dom.window.document.querySelector<HTMLElement>("#save-status")
        ?.textContent,
    ).toBe("Saved locally.");
  });

  it("reports an incomplete handoff without acknowledging it", async () => {
    delete localValues[documentSectionKey(documentId, 0)];
    await import("../src/editor/editor");

    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector<HTMLElement>("#save-status")
          ?.textContent,
      ).toBe("A document section is missing. Extract the page again."),
    );
    expect(sendMessage).not.toHaveBeenCalledWith({
      type: "ACK",
      id: documentId,
    });
  });

  it("shows figure, image, and table references under their section", async () => {
    localValues[documentSectionKey(documentId, 0)] = {
      ...section,
      html:
        "<p>See Fig. 2 and Table 1.</p>" +
        '<figure><img src="https://example.test/figure.png" alt="Figure image">' +
        "<figcaption>Fig. 2. Caption text</figcaption></figure>" +
        '<p><img src="https://example.test/photo.png" alt="Standalone image"></p>' +
        "<table><caption>Table 1. Sample data</caption><tbody><tr><td>Value</td></tr></tbody></table>",
    };
    await import("../src/editor/editor");

    await vi.waitFor(() => {
      const references = dom.window.document.querySelector(
        `[data-outline-section="${sectionId}"] .outline-references`,
      );
      expect(references?.textContent).toContain("Fig. 2");
      expect(references?.textContent).toContain("Image 1");
      expect(references?.textContent).toContain("Table 1");
      expect(
        references?.querySelector('a[href="#p2w_figure_1"]')?.textContent,
      ).toBe("Fig. 2");
    });

    expect(
      dom.window.document.querySelector(`#${sectionId} a[href="#p2w_figure_1"]`)
        ?.textContent,
    ).toBe("Fig. 2");
    expect(
      dom.window.document.querySelector(`#${sectionId} a[href="#p2w_table_1"]`)
        ?.textContent,
    ).toBe("Table 1");
  });

  it("removes original site links by default and restores them when enabled", async () => {
    localValues[documentSectionKey(documentId, 0)] = {
      ...section,
      html: '<p><a href="https://example.test/source">Original source</a></p>',
    };
    await import("../src/editor/editor");

    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelector(
          `#${sectionId} a[href="https://example.test/source"]`,
        ),
      ).toBeNull();
      expect(
        dom.window.document.querySelector(
          `#${sectionId} span[data-p2w-original-href="https://example.test/source"]`,
        )?.textContent,
      ).toBe("Original source");
    });

    storageChangedListener?.(
      {
        [DOCX_SETTINGS_KEY]: {
          newValue: {
            ...DEFAULT_DOCX_SETTINGS,
            keepOriginalLinks: true,
          },
        },
      },
      "local",
    );

    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector(
          `#${sectionId} a[href="https://example.test/source"]`,
        )?.textContent,
      ).toBe("Original source"),
    );
  });
});
