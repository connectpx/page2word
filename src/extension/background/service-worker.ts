import type { ExtractedDoc } from "../../shared/types";
import {
  documentAssetKey,
  documentMetaKey,
  documentSectionKey,
} from "../../shared/storage-keys";
import {
  DEFAULT_DOCX_SETTINGS,
  DOCX_SETTINGS_KEY,
  DOCX_SETTINGS_VERSION,
  DOCX_SETTINGS_VERSION_KEY,
  resolveStoredDocxSettings,
} from "../../shared/html-to-docx/settings";
import { captureDocumentAssets } from "./assets";

const EDITOR_URL = chrome.runtime.getURL("editor/index.html");
const BROWSER_EDITOR_URL = "http://localhost:5173/";

chrome.runtime.onMessage.addListener(
  (message: unknown, sender, sendResponse: (response: unknown) => void) => {
    if (!isRecord(message)) {
      return;
    }

    if (
      message.type === "EXTRACT" &&
      typeof message.tabId === "number" &&
      Number.isInteger(message.tabId)
    ) {
      const editorMode =
        message.editorMode === "browser" ? "browser" : "built-in";
      void extractAndOpen(message.tabId, editorMode)
        .then(({ document, openedIn }) =>
          sendResponse({
            ok: true,
            title: document.title,
            words: document.stats.words,
            openedIn,
          }),
        )
        .catch((error: unknown) =>
          sendResponse({
            ok: false,
            error:
              error instanceof Error
                ? error.message
                : "Article extraction failed.",
          }),
        );
      return true;
    }

    if (
      message.type === "ACK" &&
      typeof message.id === "string" &&
      sender.url &&
      isEditorUrl(sender.url)
    ) {
      void removeDocumentSections(message.id).catch((error: unknown) =>
        console.error("Could not clear the editor handoff payload.", error),
      );
      return;
    }

    if (
      message.type === "OPEN_OPTIONS" &&
      sender.url &&
      isEditorUrl(sender.url)
    ) {
      void chrome.runtime
        .openOptionsPage()
        .catch((error: unknown) =>
          console.error("Could not open Page2Word settings.", error),
        );
    }
  },
);

async function extractAndOpen(
  tabId: number,
  editorMode: "built-in" | "browser",
): Promise<{ document: ExtractedDoc; openedIn: "built-in" | "browser" }> {
  const editorUrl =
    editorMode === "browser" ? await requireBrowserEditor() : EDITOR_URL;
  const document = await extractFromTab(tabId);
  const capturedAssets = await captureDocumentAssets(document);
  const savedSettings = await chrome.storage.local.get([
    DOCX_SETTINGS_KEY,
    DOCX_SETTINGS_VERSION_KEY,
  ]);
  const settings = resolveStoredDocxSettings(
    savedSettings[DOCX_SETTINGS_KEY] ?? DEFAULT_DOCX_SETTINGS,
    savedSettings[DOCX_SETTINGS_VERSION_KEY],
  );
  if (savedSettings[DOCX_SETTINGS_VERSION_KEY] !== DOCX_SETTINGS_VERSION) {
    await chrome.storage.local.set({
      [DOCX_SETTINGS_KEY]: settings,
      [DOCX_SETTINGS_VERSION_KEY]: DOCX_SETTINGS_VERSION,
    });
  }
  const sections = document.contentSections;
  const storedDocument = {
    id: document.id,
    title: document.title,
    byline: document.byline,
    siteName: document.siteName,
    sourceUrl: document.sourceUrl,
    extractedAt: document.extractedAt,
    lang: document.lang,
    sectionCount: sections.length,
    sectionsAvailable: true,
    sections: sections.map(({ id, title, words }) => ({ id, title, words })),
    assets: document.assets,
    stats: document.stats,
    warnings: document.warnings,
    settings,
  };
  const payload: Record<string, unknown> = {
    [documentMetaKey(document.id)]: storedDocument,
  };
  sections.forEach((section, index) => {
    payload[documentSectionKey(document.id, index)] = section;
  });
  for (const asset of capturedAssets) {
    payload[documentAssetKey(document.id, asset.metadata.id)] = asset.base64;
  }
  await chrome.storage.local.set(payload);
  try {
    await chrome.tabs.create({
      url: `${editorUrl}#doc=${encodeURIComponent(document.id)}`,
    });
  } catch (error) {
    await chrome.storage.local.remove(Object.keys(payload));
    throw error;
  }
  return { document, openedIn: editorMode };
}

async function requireBrowserEditor(): Promise<string> {
  const unavailableMessage =
    'The Node editor server is not running. Double-click "StartServer.bat" in the Page2Word folder, then try again.';
  const response = await fetchEditorFile(
    `${BROWSER_EDITOR_URL}health`,
    unavailableMessage,
  );
  if (response.status === 404) {
    const pageResponse = await fetchEditorFile(
      BROWSER_EDITOR_URL,
      unavailableMessage,
    );
    const html = await pageResponse.text();
    if (
      pageResponse.ok &&
      html.includes("<title>Page2Word Editor</title>") &&
      html.includes('id="editor-content"')
    ) {
      return BROWSER_EDITOR_URL;
    }
    throw new Error(unavailableMessage);
  }
  if (!response.ok) {
    throw new Error(unavailableMessage);
  }
  let health: unknown;
  try {
    health = await response.json();
  } catch {
    throw new Error(unavailableMessage);
  }
  if (!isRecord(health) || health.service !== "page2word-editor") {
    throw new Error(unavailableMessage);
  }
  return BROWSER_EDITOR_URL;
}

async function fetchEditorFile(
  url: string,
  failureMessage: string,
): Promise<Response> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 1_500);
  try {
    return await fetch(url, { cache: "no-store", signal: controller.signal });
  } catch {
    throw new Error(failureMessage);
  } finally {
    clearTimeout(timeout);
  }
}

async function extractFromTab(tabId: number): Promise<ExtractedDoc> {
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["extractor/extract.js"],
  });
  const [result] = await chrome.scripting.executeScript({
    target: { tabId },
    func: () => {
      const extract = (
        window as Window & { page2wordExtract?: () => ExtractedDoc }
      ).page2wordExtract;
      if (!extract) {
        throw new Error("The article extractor could not be initialized.");
      }
      return extract();
    },
  });
  if (!result?.result) {
    throw new Error("No article content was returned from the page.");
  }
  return result.result;
}

async function removeDocumentSections(id: string): Promise<void> {
  if (!/^[a-z0-9-]+$/i.test(id)) {
    return;
  }
  const key = documentMetaKey(id);
  const stored = await chrome.storage.local.get(key);
  const metadata = stored[key];
  if (!isRecord(metadata)) {
    return;
  }
  const sectionCount = metadata.sectionCount;
  if (
    typeof sectionCount !== "number" ||
    !Number.isInteger(sectionCount) ||
    sectionCount < 0 ||
    sectionCount > 10_000
  ) {
    return;
  }
  await chrome.storage.local.set({
    [key]: { ...metadata, sectionsAvailable: false },
  });
  const keys = Array.from({ length: sectionCount }, (_, index) =>
    documentSectionKey(id, index),
  );
  const assetKeys = Array.isArray(metadata.assets)
    ? metadata.assets.flatMap((asset) =>
        isRecord(asset) && typeof asset.id === "string"
          ? [documentAssetKey(id, asset.id)]
          : [],
      )
    : [];
  await chrome.storage.local.remove([...keys, ...assetKeys]);
}

function isEditorUrl(value: string): boolean {
  try {
    const url = new URL(value);
    const editorUrl = new URL(EDITOR_URL);
    return (
      (url.origin === editorUrl.origin &&
        url.pathname === editorUrl.pathname) ||
      (url.origin === new URL(BROWSER_EDITOR_URL).origin &&
        url.pathname === new URL(BROWSER_EDITOR_URL).pathname)
    );
  } catch {
    return false;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
