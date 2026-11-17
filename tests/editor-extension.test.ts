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
    vi.stubGlobal("Element", dom.window.Element);
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

  it("updates the visible section word count when content changes", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );
    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    sectionElement!.innerHTML = "<p>Updated text has five words now.</p>";
    sectionElement!.dispatchEvent(
      new dom.window.Event("input", { bubbles: true }),
    );

    expect(sectionElement?.dataset.words).toBe("6");
    expect(
      dom.window.document.querySelector(`[data-outline-words="${sectionId}"]`)
        ?.textContent,
    ).toBe("6 words");
  });

  it("supports undo and redo for section edits", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );

    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    sectionElement!.innerHTML = "<p>Revised article text.</p>";
    sectionElement!.dispatchEvent(
      new dom.window.InputEvent("input", {
        bubbles: true,
        inputType: "insertText",
      }),
    );
    sectionElement!.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        ctrlKey: true,
        key: "z",
      }),
    );
    expect(sectionElement?.textContent).toContain("Extension page content.");

    sectionElement!.dispatchEvent(
      new dom.window.KeyboardEvent("keydown", {
        bubbles: true,
        cancelable: true,
        ctrlKey: true,
        key: "y",
      }),
    );
    expect(sectionElement?.textContent).toContain("Revised article text.");
  });

  it("restores a section to its original content", async () => {
    await import("../src/editor/editor");
    vi.spyOn(dom.window, "confirm").mockReturnValue(true);
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );

    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    sectionElement!.innerHTML = "<p>Revised article text.</p>";
    sectionElement!.dispatchEvent(
      new dom.window.Event("input", { bubbles: true }),
    );
    dom.window.document
      .querySelector<HTMLButtonElement>(`[data-restore-section="${sectionId}"]`)
      ?.click();

    await vi.waitFor(() =>
      expect(sectionElement?.textContent).toContain("Extension page content."),
    );
    expect(storageMocks.saveEditorSections).toHaveBeenCalled();
  });

  it("uses the current extraction as a restore baseline for older saved edits", async () => {
    vi.spyOn(dom.window, "confirm").mockReturnValue(true);
    storageMocks.loadEditorDocument.mockResolvedValue({
      document: {
        ...meta,
        updatedAt: "2026-11-13T00:00:00.000Z",
      },
      sections: [
        {
          documentId,
          ...section,
          html: "<p>Saved edited content.</p>",
          done: false,
        },
      ],
      assets: [],
    });
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Saved edited content."),
    );

    dom.window.document
      .querySelector<HTMLButtonElement>(`[data-restore-section="${sectionId}"]`)
      ?.click();

    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector(`[data-p2w-section="${sectionId}"]`)
          ?.textContent,
      ).toContain("Extension page content."),
    );
  });

  it("restores the whole document to its original content", async () => {
    await import("../src/editor/editor");
    vi.spyOn(dom.window, "confirm").mockReturnValue(true);
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );

    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    sectionElement!.innerHTML = "<p>Revised article text.</p>";
    sectionElement!.dispatchEvent(
      new dom.window.Event("input", { bubbles: true }),
    );
    dom.window.document
      .querySelector<HTMLButtonElement>("#restore-document")
      ?.click();

    await vi.waitFor(() =>
      expect(sectionElement?.textContent).toContain("Extension page content."),
    );
    expect(storageMocks.saveEditorSections).toHaveBeenCalled();
  });

  it("shows the selection inspector and applies editable text formatting", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );
    expect(
      dom.window.document
        .querySelector<HTMLElement>(".paper")
        ?.style.getPropertyValue("--document-font-size"),
    ).toBe("12pt");

    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    const paragraph = sectionElement?.querySelector("p");
    const range = dom.window.document.createRange();
    range.selectNodeContents(paragraph!);
    range.collapse(false);
    const selection = dom.window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    dom.window.document.dispatchEvent(new dom.window.Event("selectionchange"));

    const toolbar =
      dom.window.document.querySelector<HTMLElement>("#format-toolbar");
    expect(toolbar?.hidden).toBe(false);
    expect(toolbar?.querySelector('[data-format="font-weight"]')).toBeNull();
    range.selectNodeContents(paragraph!);
    selection?.removeAllRanges();
    selection?.addRange(range);
    dom.window.document.dispatchEvent(new dom.window.Event("selectionchange"));
    const family = toolbar?.querySelector<HTMLInputElement>(
      '[data-format="font-family"]',
    );
    const size = toolbar?.querySelector<HTMLInputElement>(
      '[data-format="font-size"]',
    );
    expect(family?.disabled).toBe(true);
    expect(size?.disabled).toBe(true);
    const color = toolbar?.querySelector<HTMLInputElement>(
      '[data-format="color"]',
    );
    color!.value = "#123456";
    color!.dispatchEvent(new dom.window.Event("change", { bubbles: true }));

    const formattedText =
      sectionElement?.querySelector<HTMLElement>("span[style]");
    expect(formattedText?.textContent).toBe("Extension page content.");
    expect(formattedText?.style.color).toContain("18, 52, 86");
  });

  it("keeps the right inspector open while clicking its controls", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );
    const paragraph = dom.window.document.querySelector(
      `[data-p2w-section="${sectionId}"] p`,
    );
    const range = dom.window.document.createRange();
    range.selectNodeContents(paragraph!);
    const selection = dom.window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    dom.window.document.dispatchEvent(new dom.window.Event("selectionchange"));

    const panel =
      dom.window.document.querySelector<HTMLElement>("#properties-panel");
    expect(panel?.hidden).toBe(false);
    panel
      ?.querySelector<HTMLInputElement>('[data-format="font-family"]')
      ?.dispatchEvent(new dom.window.Event("pointerdown", { bubbles: true }));
    selection?.removeAllRanges();
    dom.window.document.dispatchEvent(new dom.window.Event("selectionchange"));

    expect(panel?.hidden).toBe(false);
  });

  it("shows and applies table properties in the inspector", async () => {
    localValues[documentSectionKey(documentId, 0)] = {
      ...section,
      html:
        "<table><caption>Table 1. Example</caption><tbody>" +
        "<tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr>" +
        "</tbody></table>",
    };
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content table"),
      ).not.toBeNull(),
    );

    const table = dom.window.document.querySelector<HTMLTableElement>(
      "#editor-content table",
    );
    table!
      .querySelector("td")!
      .dispatchEvent(new dom.window.MouseEvent("click", { bubbles: true }));
    const panel =
      dom.window.document.querySelector<HTMLElement>("#table-properties");
    expect(panel?.hidden).toBe(false);
    expect(
      dom.window.document.querySelector<HTMLInputElement>("#table-row-count")
        ?.value,
    ).toBe("2");
    expect(
      dom.window.document.querySelector<HTMLInputElement>("#table-column-count")
        ?.value,
    ).toBe("2");
    expect(
      dom.window.document.querySelector<HTMLTextAreaElement>(
        "#table-caption-value",
      )?.value,
    ).toBe("Table 1. Example");

    const size =
      dom.window.document.querySelector<HTMLInputElement>("#table-font-size");
    size!.value = "14";
    size!.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    const alignment = dom.window.document.querySelector<HTMLSelectElement>(
      "#table-cell-alignment",
    );
    alignment!.value = "center";
    alignment!.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    const padding = dom.window.document.querySelector<HTMLInputElement>(
      "#table-cell-padding",
    );
    padding!.value = "4";
    padding!.dispatchEvent(new dom.window.Event("change", { bubbles: true }));
    const borders =
      dom.window.document.querySelector<HTMLSelectElement>(
        "#table-border-mode",
      );
    borders!.value = "grid";
    borders!.dispatchEvent(new dom.window.Event("change", { bubbles: true }));

    const cells = table!.querySelectorAll<HTMLTableCellElement>("th,td");
    expect(cells[0].style.fontSize).toBe("14pt");
    expect(cells[1].style.textAlign).toBe("center");
    expect(cells[2].style.padding).toBe("4px");
    expect(cells[3].style.border).toContain("solid");
  });

  it("shows image actions in a compact menu and image properties on selection", async () => {
    localValues[documentSectionKey(documentId, 0)] = {
      ...section,
      html:
        '<figure><img src="https://example.test/figure.png" alt="Example image">' +
        "<figcaption>Figure 1. Example</figcaption></figure>",
    };
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("[data-image-menu-toggle]"),
      ).not.toBeNull(),
    );

    dom.window.document
      .querySelector<HTMLButtonElement>("[data-image-menu-toggle]")
      ?.click();

    expect(
      dom.window.document.querySelector<HTMLElement>(".image-edit-actions")
        ?.hidden,
    ).toBe(false);
    expect(
      dom.window.document.querySelector<HTMLElement>("#image-properties")
        ?.hidden,
    ).toBe(false);
    expect(
      dom.window.document.querySelector<HTMLTextAreaElement>("#image-alt-value")
        ?.value,
    ).toBe("Example image");
  });

  it("pastes normalized plain text and updates its word count", async () => {
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );
    const sectionElement = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    const paragraph = sectionElement?.querySelector("p");
    paragraph!.append(dom.window.document.createTextNode(" "));
    const range = dom.window.document.createRange();
    range.selectNodeContents(paragraph!);
    range.collapse(false);
    const selection = dom.window.document.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    const paste = new dom.window.Event("paste", {
      bubbles: true,
      cancelable: true,
    });
    Object.defineProperty(paste, "clipboardData", {
      value: {
        getData: () => "New—text\nsecond\uFFFD line",
      },
    });
    paragraph!.dispatchEvent(paste);

    expect(paste.defaultPrevented).toBe(true);
    expect(sectionElement?.textContent).toContain("New-text");
    expect(sectionElement?.textContent).not.toContain("\uFFFD");
    expect(
      dom.window.document.querySelector(`[data-outline-words="${sectionId}"]`)
        ?.textContent,
    ).toBe("6 words");
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

  it("shows navigable figures and tables in a separate sidebar section", async () => {
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
      expect(
        dom.window.document.querySelector("#figure-outline")?.textContent,
      ).toBe("Figure 2");
      expect(
        dom.window.document.querySelector("#table-outline")?.textContent,
      ).toBe("Table 1");
    });

    expect(
      dom.window.document.querySelector(
        `[data-outline-section="${sectionId}"] .outline-references`,
      ),
    ).toBeNull();
    expect(
      dom.window.document.querySelector(`#${sectionId} a[href="#p2w_figure_1"]`)
        ?.textContent,
    ).toBe("Figure 2");
    expect(
      dom.window.document.querySelector(`#${sectionId} a[href="#p2w_table_1"]`)
        ?.textContent,
    ).toBe("Table 1");

    const sectionTarget = dom.window.document.getElementById(sectionId);
    const figureTarget = dom.window.document.getElementById("p2w_figure_1");
    const sectionScroll = vi.fn();
    const figureScroll = vi.fn();
    Object.defineProperty(sectionTarget, "scrollIntoView", {
      value: sectionScroll,
    });
    Object.defineProperty(figureTarget, "scrollIntoView", {
      value: figureScroll,
    });
    const sectionClick = new dom.window.MouseEvent("click", {
      bubbles: true,
      cancelable: true,
    });
    dom.window.document
      .querySelector<HTMLAnchorElement>(
        `[data-outline-section-link="${sectionId}"]`,
      )
      ?.dispatchEvent(sectionClick);
    expect(sectionClick.defaultPrevented).toBe(true);
    expect(dom.window.location.hash).toBe(`#doc=${documentId}`);
    expect(sectionScroll).toHaveBeenCalledOnce();

    const figureClick = new dom.window.MouseEvent("click", {
      bubbles: true,
      cancelable: true,
    });
    dom.window.document
      .querySelector<HTMLAnchorElement>("#figure-outline a")
      ?.dispatchEvent(figureClick);
    expect(figureClick.defaultPrevented).toBe(true);
    expect(dom.window.location.hash).toBe(`#doc=${documentId}`);
    expect(figureScroll).toHaveBeenCalledOnce();
  });

  it("excludes matching section titles and restores them when keywords change", async () => {
    const title =
      "Declaration of Generative AI and AI-assisted technologies in writing";
    localValues[documentMetaKey(documentId)] = {
      ...meta,
      sections: [{ id: sectionId, title, words: 2 }],
      settings: {
        excludedSectionKeywords: ["Declaration of Generative AI"],
      },
    };
    localValues[documentSectionKey(documentId, 0)] = {
      ...section,
      title,
    };
    await import("../src/editor/editor");

    const editorSection = dom.window.document.querySelector<HTMLElement>(
      `[data-p2w-section="${sectionId}"]`,
    );
    const outlineItem = dom.window.document.querySelector<HTMLElement>(
      `[data-outline-section="${sectionId}"]`,
    );
    await vi.waitFor(() => {
      expect(editorSection?.hidden).toBe(true);
      expect(outlineItem?.hidden).toBe(true);
    });

    storageChangedListener?.(
      {
        [DOCX_SETTINGS_KEY]: {
          newValue: {
            ...DEFAULT_DOCX_SETTINGS,
            excludedSectionKeywords: [],
          },
        },
      },
      "local",
    );
    await vi.waitFor(() => {
      expect(editorSection?.hidden).toBe(false);
      expect(outlineItem?.hidden).toBe(false);
    });
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

  it("imports Word-style HTML into heading sections with an unlinked bold table caption", async () => {
    vi.stubGlobal("DOMParser", dom.window.DOMParser);
    vi.stubGlobal("history", dom.window.history);
    vi.stubGlobal("crypto", { randomUUID: () => "imported-test" });
    await import("../src/editor/editor");
    await vi.waitFor(() =>
      expect(
        dom.window.document.querySelector("#editor-content")?.textContent,
      ).toContain("Extension page content."),
    );

    const pasteInput =
      dom.window.document.querySelector<HTMLTextAreaElement>("#paste-html");
    pasteInput!.value =
      "<html><head><title>Word paper</title></head><body>" +
      "<div class=WordSection1><p>Document title</p>" +
      "<h1>Abstract</h1><p>Summary text.</p>" +
      "<h1>1. Introduction</h1><p>Introduction text.</p>" +
      "<h2>1.1. Background</h2><p>Background text.</p>" +
      "<p><b>Table 1.</b> Summary of key findings.</p>" +
      "<div align=center><table><tr><td>Finding</td></tr>" +
      "<tr><td>Result</td></tr></table></div>" +
      "<h1>2. Methodology</h1><p>Methods text.</p>" +
      "</div></body></html>";
    dom.window.document
      .querySelector<HTMLButtonElement>("#load-paste")
      ?.click();

    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelectorAll("[data-p2w-section]"),
      ).toHaveLength(3);
    });

    const importedSections = [
      ...dom.window.document.querySelectorAll<HTMLElement>(
        "[data-p2w-section]",
      ),
    ];
    expect(
      importedSections.map((item) => item.getAttribute("aria-label")),
    ).toEqual(["Abstract", "1. Introduction", "2. Methodology"]);
    expect(importedSections[1].querySelector("h3")?.textContent).toBe(
      "1.1. Background",
    );
    expect(
      importedSections[1].querySelector(".p2w-table-caption b")?.textContent,
    ).toBe("Table 1.");
    expect(
      importedSections[1].querySelector(".p2w-table-caption a"),
    ).toBeNull();
    expect(
      importedSections[1].querySelector("table tr:first-child th")?.textContent,
    ).toBe("Finding");
    expect(storageMocks.saveEditorDocument).toHaveBeenLastCalledWith(
      expect.objectContaining({ sectionCount: 3 }),
      expect.any(Array),
    );
  });
});
