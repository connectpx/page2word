import { beforeEach, describe, expect, it, vi } from "vitest";
import type { ExtractedDoc } from "../src/shared/types";
import {
  documentAssetKey,
  documentMetaKey,
  documentSectionKey,
} from "../src/shared/storage-keys";
import {
  DEFAULT_DOCX_SETTINGS,
  DOCX_SETTINGS_KEY,
  DOCX_SETTINGS_VERSION_KEY,
} from "../src/shared/html-to-docx/settings";

type MessageListener = (
  message: unknown,
  sender: { url?: string },
  sendResponse: (response: unknown) => void,
) => boolean | void;

const extractedDocument: ExtractedDoc = {
  id: "document-test-123",
  title: "Test article",
  sourceUrl: "https://example.test/article",
  extractedAt: "2026-11-12T00:00:00.000Z",
  html: "<p>Example article content.</p>",
  sections: [{ id: "section-1", title: "Test article", words: 3 }],
  contentSections: [
    {
      id: "section-1",
      title: "Test article",
      words: 3,
      html: "<p>Example article content.</p>",
    },
  ],
  assets: [],
  stats: { words: 3, images: 0, tables: 0, charts: 0 },
  warnings: [],
};

describe("editor handoff service worker", () => {
  let listener: MessageListener | undefined;
  let storedValues: Map<string, unknown>;
  let chromeMock: {
    runtime: {
      getURL: (path: string) => string;
      onMessage: { addListener: (callback: MessageListener) => void };
    };
    scripting: {
      executeScript: ReturnType<typeof vi.fn>;
    };
    storage: {
      local: {
        set: ReturnType<typeof vi.fn>;
        get: ReturnType<typeof vi.fn>;
        remove: ReturnType<typeof vi.fn>;
      };
    };
    tabs: { create: ReturnType<typeof vi.fn> };
  };

  beforeEach(async () => {
    vi.resetModules();
    vi.unstubAllGlobals();
    storedValues = new Map();
    listener = undefined;

    chromeMock = {
      runtime: {
        getURL: vi.fn((path: string) => `chrome-extension://test-id/${path}`),
        onMessage: {
          addListener: (callback) => {
            listener = callback;
          },
        },
      },
      scripting: {
        executeScript: vi.fn(async (details: { files?: string[] }) =>
          details.files ? [] : [{ result: extractedDocument }],
        ),
      },
      storage: {
        local: {
          set: vi.fn(async (values: Record<string, unknown>) => {
            for (const [key, value] of Object.entries(values)) {
              storedValues.set(key, value);
            }
          }),
          get: vi.fn(async (key: string | string[]) => {
            const keys = Array.isArray(key) ? key : [key];
            return Object.fromEntries(
              keys
                .filter((storedKey) => storedValues.has(storedKey))
                .map((storedKey) => [storedKey, storedValues.get(storedKey)]),
            );
          }),
          remove: vi.fn(async (keys: string[]) => {
            for (const key of keys) {
              storedValues.delete(key);
            }
          }),
        },
      },
      tabs: { create: vi.fn(async () => ({ id: 21 })) },
    };
    vi.stubGlobal("chrome", chromeMock);
    await import("../src/extension/background/service-worker");
  });

  it("stores sections locally and opens the editor with only the document ID", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const response = vi.fn();
    expect(listener({ type: "EXTRACT", tabId: 7 }, {}, response)).toBe(true);
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(response).toHaveBeenCalledWith({
      ok: true,
      title: extractedDocument.title,
      words: extractedDocument.stats.words,
      openedIn: "built-in",
    });

    expect(chromeMock.scripting.executeScript).toHaveBeenCalledTimes(2);
    expect(
      storedValues.get(documentSectionKey(extractedDocument.id, 0)),
    ).toEqual(extractedDocument.contentSections[0]);
    expect(
      storedValues.get(documentMetaKey(extractedDocument.id)),
    ).toMatchObject({
      settings: DEFAULT_DOCX_SETTINGS,
    });
    const createdUrl = chromeMock.tabs.create.mock.calls[0]?.[0].url as string;
    expect(createdUrl).toContain(
      `chrome-extension://test-id/editor/index.html#doc=${extractedDocument.id}`,
    );
    expect(createdUrl).not.toContain(extractedDocument.html);
  });

  it("migrates a saved 13-point subheading default during extraction", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    storedValues.set(DOCX_SETTINGS_KEY, { subheadingFontSizePt: 13 });
    const response = vi.fn();

    expect(listener({ type: "EXTRACT", tabId: 7 }, {}, response)).toBe(true);
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(storedValues.get(DOCX_SETTINGS_KEY)).toMatchObject({
      subheadingFontSizePt: 12,
    });
    expect(storedValues.get(DOCX_SETTINGS_VERSION_KEY)).toBe(4);
    expect(
      storedValues.get(documentMetaKey(extractedDocument.id)),
    ).toMatchObject({
      settings: { subheadingFontSizePt: 12 },
    });
  });

  it("opens the browser editor when the Node server is available", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const fetchMock = vi.fn(
      async () =>
        new Response(JSON.stringify({ service: "page2word-editor" }), {
          status: 200,
          headers: { "content-type": "application/json" },
        }),
    );
    vi.stubGlobal("fetch", fetchMock);
    const response = vi.fn();

    listener(
      { type: "EXTRACT", tabId: 7, editorMode: "browser" },
      {},
      response,
    );
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(fetchMock).toHaveBeenCalledWith(
      "http://localhost:5173/health",
      expect.objectContaining({ cache: "no-store" }),
    );
    expect(response).toHaveBeenCalledWith(
      expect.objectContaining({ ok: true, openedIn: "browser" }),
    );
    expect(chromeMock.tabs.create).toHaveBeenCalledWith({
      url: `http://localhost:5173/#doc=${extractedDocument.id}`,
    });
  });

  it("supports the Node server when its health endpoint is unavailable", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(new Response("", { status: 404 }))
      .mockResolvedValueOnce(
        new Response(
          '<html><head><title>Page2Word Editor</title></head><body><main id="editor-content"></main></body></html>',
          { status: 200, headers: { "content-type": "text/html" } },
        ),
      );
    vi.stubGlobal("fetch", fetchMock);
    const response = vi.fn();

    listener(
      { type: "EXTRACT", tabId: 7, editorMode: "browser" },
      {},
      response,
    );
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(fetchMock).toHaveBeenNthCalledWith(
      1,
      "http://localhost:5173/health",
      expect.objectContaining({ cache: "no-store" }),
    );
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      "http://localhost:5173/",
      expect.objectContaining({ cache: "no-store" }),
    );
    expect(response).toHaveBeenCalledWith(
      expect.objectContaining({ ok: true, openedIn: "browser" }),
    );
  });

  it("rejects a different service running on the Node editor port", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const fetchMock = vi.fn(
      async () =>
        new Response(JSON.stringify({ service: "other-service" }), {
          status: 200,
          headers: { "content-type": "application/json" },
        }),
    );
    vi.stubGlobal("fetch", fetchMock);
    const response = vi.fn();

    listener(
      { type: "EXTRACT", tabId: 7, editorMode: "browser" },
      {},
      response,
    );
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(fetchMock).toHaveBeenCalledOnce();
    expect(response).toHaveBeenCalledWith({
      ok: false,
      error:
        'The Node editor server is not running. Double-click "StartServer.bat" in the Page2Word folder, then try again.',
    });
    expect(chromeMock.scripting.executeScript).not.toHaveBeenCalled();
  });

  it("explains how to start the Node editor server when it is unavailable", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        throw new TypeError("Failed to fetch");
      }),
    );
    const response = vi.fn();

    listener(
      { type: "EXTRACT", tabId: 7, editorMode: "browser" },
      {},
      response,
    );
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    expect(response).toHaveBeenCalledWith({
      ok: false,
      error:
        'The Node editor server is not running. Double-click "StartServer.bat" in the Page2Word folder, then try again.',
    });
    expect(chromeMock.scripting.executeScript).not.toHaveBeenCalled();
    expect(chromeMock.tabs.create).not.toHaveBeenCalled();
  });

  it("clears handoff sections only after an ACK from the editor origin", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const response = vi.fn();
    listener({ type: "EXTRACT", tabId: 7 }, {}, response);
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    listener(
      { type: "ACK", id: extractedDocument.id },
      { url: "https://untrusted.example/editor" },
      vi.fn(),
    );
    expect(storedValues.has(documentSectionKey(extractedDocument.id, 0))).toBe(
      true,
    );

    listener(
      { type: "ACK", id: extractedDocument.id },
      {
        url: `chrome-extension://test-id/editor/index.html#doc=${extractedDocument.id}`,
      },
      vi.fn(),
    );
    await vi.waitFor(() =>
      expect(
        storedValues.has(documentSectionKey(extractedDocument.id, 0)),
      ).toBe(false),
    );
    expect(
      storedValues.get(documentMetaKey(extractedDocument.id)),
    ).toMatchObject({
      sectionsAvailable: false,
    });
  });

  it("stores captured image metadata and bytes with the editor handoff", async () => {
    if (!listener) {
      throw new Error(
        "The service worker did not register its message listener.",
      );
    }
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j6sgAAAAASUVORK5CYII=",
      "base64",
    );
    const imageDocument: ExtractedDoc = {
      ...extractedDocument,
      contentSections: [
        {
          ...extractedDocument.contentSections[0],
          html: '<p><img src="https://example.test/chart.png" alt="Chart" width="1" height="1"></p>',
        },
      ],
      assets: [],
      stats: { ...extractedDocument.stats, images: 1 },
    };
    chromeMock.scripting.executeScript.mockImplementation(
      async (details: { files?: string[] }) =>
        details.files ? [] : [{ result: imageDocument }],
    );
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(png, {
            status: 200,
            headers: { "content-type": "image/png" },
          }),
      ),
    );
    const response = vi.fn();

    listener({ type: "EXTRACT", tabId: 7 }, {}, response);
    await vi.waitFor(() => expect(response).toHaveBeenCalledOnce());

    const metadata = storedValues.get(documentMetaKey(imageDocument.id)) as {
      assets: { id: string; bytes: number }[];
    };
    expect(metadata.assets).toHaveLength(1);
    expect(metadata.assets[0].bytes).toBe(png.byteLength);
    const assetKey = documentAssetKey(imageDocument.id, metadata.assets[0].id);
    expect(storedValues.get(assetKey)).toBe(png.toString("base64"));
    expect(
      storedValues.get(documentSectionKey(imageDocument.id, 0)),
    ).toMatchObject({
      html: expect.stringContaining(`p2w-asset:${metadata.assets[0].id}`),
    });
    listener(
      { type: "ACK", id: imageDocument.id },
      { url: `http://localhost:5173/#doc=${imageDocument.id}` },
      vi.fn(),
    );
    await vi.waitFor(() => expect(storedValues.has(assetKey)).toBe(false));
  });
});
