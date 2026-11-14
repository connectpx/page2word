import { sanitizeHtml } from "../shared/sanitize";
import { normalizeArticleMarkup } from "../shared/normalize-article-markup";
import {
  DEFAULT_DOCX_SETTINGS,
  DOCX_SETTINGS_KEY,
  resolveDocxSettings,
} from "../shared/html-to-docx/settings";
import { parseHtmlToDocxBlocks } from "../shared/html-to-docx/parse";
import type {
  DocxBlock,
  DocxImageAsset,
  DocxMetadata,
  DocxSettings,
  ExportProgress,
} from "../shared/html-to-docx/types";
import type {
  ContentSection,
  ExtractedAsset,
  StoredDocument,
} from "../shared/types";
import {
  documentAssetKey,
  documentMetaKey,
  documentSectionKey,
} from "../shared/storage-keys";
import {
  deleteEditorDocument,
  loadEditorDocument,
  saveEditorDocument,
  saveEditorSections,
  toEditorSection,
  type EditorAssetRecord,
  type EditorDocumentRecord,
  type EditorSectionRecord,
} from "./storage";

const editorRoot = document.querySelector<HTMLElement>("#editor-content");
const outline = document.querySelector<HTMLElement>("#outline-list");
const status = document.querySelector<HTMLElement>("#save-status");
const warningBanner = document.querySelector<HTMLElement>("#warning-banner");
const bridgeNotice = document.querySelector<HTMLElement>("#bridge-notice");
const fileModeNotice = document.querySelector<HTMLElement>("#file-mode-notice");
const fileInput = document.querySelector<HTMLInputElement>("#html-file");
const imageFileInput = document.querySelector<HTMLInputElement>("#image-file");
const openFileButton = document.querySelector<HTMLButtonElement>("#open-html");
const openSettingsButton =
  document.querySelector<HTMLButtonElement>("#open-settings");
const saveHtmlButton = document.querySelector<HTMLButtonElement>("#save-html");
const exportButton = document.querySelector<HTMLButtonElement>("#export-word");
const formatToolbar = document.querySelector<HTMLElement>("#format-toolbar");
const pasteInput = document.querySelector<HTMLTextAreaElement>("#paste-html");
const loadPasteButton =
  document.querySelector<HTMLButtonElement>("#load-paste");
const exportDialog =
  document.querySelector<HTMLDialogElement>("#export-dialog");
const exportProgress =
  document.querySelector<HTMLProgressElement>("#export-progress");
const exportStage = document.querySelector<HTMLElement>("#export-stage");
const exportSummary = document.querySelector<HTMLElement>("#export-summary");
const cancelExportButton =
  document.querySelector<HTMLButtonElement>("#cancel-export");
const closeExportButton =
  document.querySelector<HTMLButtonElement>("#close-export");

let currentDocument: EditorDocumentRecord | undefined;
let currentAssets = new Map<string, EditorAssetRecord>();
const assetObjectUrls = new Map<string, string>();
let activeExportController: AbortController | undefined;
let activeExportWorker: Worker | undefined;
let expectedSections = 0;
let expectedSectionIds = new Set<string>();
let expectedAssetIds = new Set<string>();
let receivedSections = new Set<string>();
let receivedAssets = new Map<string, EditorAssetRecord>();
let assetErrors = new Map<string, string>();
let restoredSections = new Map<string, EditorSectionRecord>();
let dirtySections = new Set<string>();
let queuedMessages = Promise.resolve();
let saveTimer: number | undefined;
const extensionRuntimeAvailable =
  typeof chrome !== "undefined" && Boolean(chrome.runtime?.id);

if (fileModeNotice && location.protocol === "file:") {
  fileModeNotice.hidden = false;
}

function updateFormattingToolbar(): void {
  if (!formatToolbar) {
    return;
  }
  const selection = window.getSelection();
  const section = selectedEditorSection(selection);
  if (!selection || !section || selection.rangeCount === 0) {
    formatToolbar.hidden = true;
    return;
  }
  const range = selection.getRangeAt(0);
  const bounds = range.getBoundingClientRect();
  if (!bounds.width && !bounds.height) {
    formatToolbar.hidden = true;
    return;
  }
  formatToolbar.hidden = false;
  const toolbarBounds = formatToolbar.getBoundingClientRect();
  const left = Math.max(
    8,
    Math.min(
      window.innerWidth - toolbarBounds.width - 8,
      bounds.left + bounds.width / 2 - toolbarBounds.width / 2,
    ),
  );
  const top = Math.max(8, bounds.top - toolbarBounds.height - 8);
  formatToolbar.style.left = `${left}px`;
  formatToolbar.style.top = `${top}px`;
}

window.addEventListener("message", (event: MessageEvent<unknown>) => {
  if (
    event.source !== window ||
    event.origin !== location.origin ||
    !isRecord(event.data)
  ) {
    return;
  }
  const message = event.data;
  if (message.source !== "page2word") {
    return;
  }
  queuedMessages = queuedMessages
    .then(() => handleBridgeMessage(message))
    .catch((error: unknown) => showError(error));
});

if (!extensionRuntimeAvailable) {
  window.postMessage(
    { source: "page2word-editor", type: "READY" },
    location.origin,
  );
}
if (extensionRuntimeAvailable) {
  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName === "local" && changes[DOCX_SETTINGS_KEY]) {
      void updateDocumentSettings(changes[DOCX_SETTINGS_KEY].newValue).catch(
        showError,
      );
    }
  });
  const documentId = new URLSearchParams(location.hash.slice(1)).get("doc");
  if (documentId) {
    void loadExtensionDocument(documentId).catch(showError);
  } else if (bridgeNotice) {
    bridgeNotice.hidden = false;
    setStatus("Ready to open an article.");
  }
} else {
  window.setTimeout(() => {
    if (!currentDocument) {
      const id = new URLSearchParams(location.hash.slice(1)).get("doc");
      if (id) {
        void restoreSavedDocument(id).catch((error: unknown) =>
          showError(error),
        );
      } else if (bridgeNotice) {
        bridgeNotice.hidden = false;
      }
    }
  }, 1800);
}

openFileButton?.addEventListener("click", () => fileInput?.click());
fileInput?.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) {
    void loadHtmlFile(file);
  }
  fileInput.value = "";
});
loadPasteButton?.addEventListener("click", () => {
  if (pasteInput?.value.trim()) {
    void loadHtmlContent(pasteInput.value, "Pasted article");
  }
});

imageFileInput?.addEventListener("change", () => {
  const file = imageFileInput.files?.[0];
  const assetId = imageFileInput.dataset.assetId;
  const imageKey = imageFileInput.dataset.imageKey;
  imageFileInput.value = "";
  delete imageFileInput.dataset.assetId;
  delete imageFileInput.dataset.imageKey;
  if (file && imageKey) {
    void replaceImageAsset(assetId, imageKey, file).catch((error: unknown) =>
      showError(error),
    );
  }
});

openSettingsButton?.addEventListener("click", () => {
  if (extensionRuntimeAvailable) {
    void chrome.runtime.openOptionsPage().catch(showError);
  } else {
    window.postMessage(
      { source: "page2word-editor", type: "OPEN_OPTIONS" },
      location.origin,
    );
  }
});

formatToolbar?.addEventListener("mousedown", (event) => {
  if (
    event.target instanceof Element &&
    event.target.closest("button[data-command]")
  ) {
    event.preventDefault();
  }
});
formatToolbar?.addEventListener("click", (event) => {
  if (!(event.target instanceof Element)) {
    return;
  }
  const button = event.target.closest<HTMLButtonElement>(
    "button[data-command]",
  );
  const command = button?.dataset.command;
  if (command !== "bold" && command !== "italic" && command !== "underline") {
    return;
  }
  const section = selectedEditorSection(window.getSelection());
  if (!section) {
    setStatus("Select text in one article section to format it.");
    return;
  }
  if (!document.execCommand(command)) {
    setStatus("The selected text could not be formatted.");
    return;
  }
  markDirty(section.dataset.p2wSection ?? "");
  updateFormattingToolbar();
});

document.addEventListener("selectionchange", updateFormattingToolbar);
document.addEventListener("keyup", updateFormattingToolbar);
document.addEventListener("mouseup", updateFormattingToolbar);
window.addEventListener("scroll", updateFormattingToolbar, true);
window.addEventListener("resize", updateFormattingToolbar);

saveHtmlButton?.addEventListener("click", () => {
  void saveStandaloneHtml().catch((error: unknown) => showError(error));
});

async function saveStandaloneHtml(): Promise<void> {
  if (!editorRoot) {
    return;
  }
  const sections = [
    ...editorRoot.querySelectorAll<HTMLElement>("[data-p2w-section]"),
  ];
  const htmlSections: string[] = [];
  for (const section of sections) {
    if (section.hidden) {
      continue;
    }
    const copy = section.cloneNode(true) as HTMLElement;
    removeEditorUi(copy);
    normalizeArticleMarkup(copy, document);
    for (const image of copy.querySelectorAll<HTMLImageElement>(
      "img[data-p2w-asset]",
    )) {
      const asset = currentAssets.get(image.dataset.p2wAsset ?? "");
      if (!asset) {
        const placeholder = document.createElement("span");
        placeholder.textContent = `[Image unavailable: ${image.alt || "image"}]`;
        image.replaceWith(placeholder);
        continue;
      }
      const bytes = new Uint8Array(await asset.blob.arrayBuffer());
      image.setAttribute(
        "src",
        `data:${asset.mime};base64,${bytesToBase64(bytes)}`,
      );
      image.removeAttribute("data-p2w-asset");
    }
    htmlSections.push(sanitizeHtml(copy.innerHTML));
  }
  const body = htmlSections.join("\n");
  const resolvedSettings = resolveDocxSettings(currentDocument?.settings);
  const settings = JSON.stringify(resolvedSettings).replace(/</g, "\\u003c");
  const title = escapeHtml(currentDocument?.title ?? document.title);
  const titleHtml = currentDocument?.title
    ? `<h1 class="p2w-document-title">${escapeHtml(currentDocument.title)}</h1>`
    : "";
  const standalone = `<!doctype html>
<html lang="${escapeHtml(currentDocument?.lang ?? "en")}">
<head><meta charset="utf-8"><title>${title}</title>
<style>
body{font:${resolvedSettings.fontSizePt}px/${resolvedSettings.lineSpacing} "${escapeHtml(resolvedSettings.fontFamily)}",serif;margin:3rem auto;max-width:48rem;padding:0 1rem;text-align:${resolvedSettings.alignment}}
h1{font-size:${resolvedSettings.headingFontSizePt}px;font-weight:bold}
h2,h3,h4,h5,h6{font-size:${resolvedSettings.subheadingFontSizePt}px;font-weight:bold}
h1.p2w-document-title{font-size:${resolvedSettings.titleFontSizePt}px;margin:0 0 1em;text-align:center}
figure{margin:20px auto 16px}
img{max-width:100%;height:auto;margin-top:8px}
table{border-collapse:collapse;max-width:100%;width:100%}
td,th{border:0;font-size:${resolvedSettings.tableFontSizePt}px;padding:.4rem;text-align:left}
thead tr>*,table tr:first-child>th{border-bottom:1px solid #202124;border-left:0;border-right:0;border-top:1px solid #202124}
figure figcaption{border-bottom:1px solid #202124;display:block;font-size:${resolvedSettings.fontSizePt}px;margin:8px 0 16px;padding-bottom:6px;text-align:left}
table caption{font-size:${resolvedSettings.fontSizePt}px;text-align:left}
</style>
<script type="application/json" id="p2w-settings">${settings}</script></head>
<body>${titleHtml}${body}</body></html>`;
  download(
    standalone,
    `${safeFilename(currentDocument?.title ?? "page2word")}.html`,
    "text/html",
  );
}

exportButton?.addEventListener("click", () => void exportWordDocument());
cancelExportButton?.addEventListener("click", () => {
  activeExportController?.abort();
  activeExportWorker?.terminate();
});
closeExportButton?.addEventListener("click", () => exportDialog?.close());

editorRoot?.addEventListener("input", (event) => {
  const section = findEditorSection(event.target);
  if (!section) {
    return;
  }
  updateOutlineReferences(section);
  markDirty(section.dataset.p2wSection ?? "");
});

editorRoot?.addEventListener("click", (event) => {
  if (!(event.target instanceof Element)) {
    return;
  }
  const button = event.target.closest<HTMLButtonElement>(
    "button[data-image-action]",
  );
  if (!button) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  if (button.dataset.imageAction === "download") {
    const assetId = button.dataset.assetId;
    const name = button.dataset.imageName || "article-image";
    void downloadArticleImage(
      assetId,
      button.dataset.imageKey ?? "",
      name,
    ).catch((error: unknown) => showError(error));
  } else if (button.dataset.imageAction === "replace" && imageFileInput) {
    imageFileInput.dataset.assetId = button.dataset.assetId ?? "";
    imageFileInput.dataset.imageKey = button.dataset.imageKey ?? "";
    imageFileInput.click();
  }
});

outline?.addEventListener("change", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") {
    return;
  }
  const sectionId = target.dataset.doneSection;
  const section = sectionId ? findSection(sectionId) : null;
  if (section) {
    markDirty(section.dataset.p2wSection ?? "");
  }
});

outline?.addEventListener("click", (event) => {
  if (!(event.target instanceof Element)) {
    return;
  }
  const button = event.target.closest<HTMLButtonElement>(
    "button[data-section-visibility]",
  );
  const sectionId = button?.dataset.sectionVisibility;
  const section = sectionId ? findSection(sectionId) : null;
  if (!button || !section || !sectionId) {
    return;
  }
  event.preventDefault();
  section.hidden = !section.hidden;
  updateSectionVisibilityButton(sectionId, section.hidden);
  markDirty(sectionId);
});

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") {
    void saveDirtySections();
  }
});
window.addEventListener("pagehide", () => {
  void saveDirtySections();
  revokeAssetObjectUrls();
});
window.addEventListener("pageshow", () => {
  initializeAssetObjectUrls(currentAssets);
});
window.addEventListener("beforeunload", (event) => {
  if (dirtySections.size > 0) {
    event.preventDefault();
    event.returnValue = "";
  }
});

async function handleBridgeMessage(
  message: Record<string, unknown>,
): Promise<void> {
  switch (message.type) {
    case "ERROR":
      if (bridgeNotice) {
        bridgeNotice.hidden = false;
        const notice = bridgeNotice.querySelector("p");
        if (notice) {
          notice.textContent =
            typeof message.error === "string"
              ? message.error
              : "The extension connection failed.";
        }
      }
      return;
    case "REOPEN":
      if (typeof message.id === "string") {
        await restoreSavedDocument(message.id);
      }
      return;
    case "DOC_BEGIN":
      await beginDocument(message);
      return;
    case "SECTION":
      await receiveSection(message);
      return;
    case "ASSET":
      await receiveAsset(message);
      return;
    case "ASSET_ERROR":
      receiveAssetError(message);
      return;
    case "DOC_END":
      await finishDocument(message);
      return;
    case "SETTINGS":
      await updateDocumentSettings(message.settings);
      return;
  }
}

async function loadExtensionDocument(id: string): Promise<void> {
  if (!/^[a-z0-9-]+$/i.test(id)) {
    throw new Error("The editor document ID is invalid.");
  }
  const metaKey = documentMetaKey(id);
  const storedMeta = await chrome.storage.local.get(metaKey);
  const meta = storedMeta[metaKey];
  if (
    !isStoredDocument(meta) ||
    meta.id !== id ||
    meta.sectionCount > 10_000 ||
    meta.sections.length !== meta.sectionCount
  ) {
    throw new Error("The extracted document has expired. Extract it again.");
  }
  if (meta.sectionsAvailable === false) {
    await restoreSavedDocument(id);
    return;
  }

  await handleBridgeMessage({
    type: "DOC_BEGIN",
    id,
    meta,
    sections: meta.sections,
  });
  const sectionKeys = Array.from({ length: meta.sectionCount }, (_, index) =>
    documentSectionKey(id, index),
  );
  for (let offset = 0; offset < sectionKeys.length; offset += 100) {
    const keys = sectionKeys.slice(offset, offset + 100);
    const storedSections = await chrome.storage.local.get(keys);
    for (const key of keys) {
      const section = storedSections[key];
      if (
        !isRecord(section) ||
        typeof section.id !== "string" ||
        typeof section.title !== "string" ||
        typeof section.words !== "number" ||
        typeof section.html !== "string"
      ) {
        throw new Error(
          "A document section is missing. Extract the page again.",
        );
      }
      await handleBridgeMessage({ type: "SECTION", id, section });
    }
  }

  for (const asset of meta.assets) {
    const key = documentAssetKey(id, asset.id);
    const storedAsset = await chrome.storage.local.get(key);
    const base64 = storedAsset[key];
    if (typeof base64 !== "string") {
      await handleBridgeMessage({
        type: "ASSET_ERROR",
        id,
        assetId: asset.id,
        error: "The image is missing from temporary extension storage.",
      });
      continue;
    }
    await handleBridgeMessage({
      type: "ASSET",
      id,
      asset: { ...asset, base64 },
    });
  }

  await handleBridgeMessage({
    type: "DOC_END",
    id,
    counts: { sections: meta.sectionCount, assets: meta.assets.length },
  });
}

async function updateDocumentSettings(value: unknown): Promise<void> {
  if (!currentDocument) {
    return;
  }
  const previousSettings = resolveDocxSettings(currentDocument.settings);
  const settings = resolveDocxSettings(value);
  currentDocument.settings = { ...settings };
  applyDocumentSettings(settings);
  const sections =
    previousSettings.keepOriginalLinks === settings.keepOriginalLinks
      ? []
      : [
          ...(editorRoot?.querySelectorAll<HTMLElement>("[data-p2w-section]") ??
            []),
        ].flatMap((section) => {
          applyOriginalLinkSetting(section, settings.keepOriginalLinks);
          const record = sectionRecord(section);
          return record ? [record] : [];
        });
  await saveEditorDocument(currentDocument, sections, [
    ...currentAssets.values(),
  ]);
}

function applyDocumentSettings(settings: DocxSettings): void {
  const paper = document.querySelector<HTMLElement>(".paper");
  if (!paper) {
    return;
  }
  paper.style.setProperty("--document-font-family", `"${settings.fontFamily}"`);
  paper.style.setProperty("--document-font-size", `${settings.fontSizePt}px`);
  paper.style.setProperty(
    "--document-line-height",
    String(settings.lineSpacing),
  );
  paper.style.setProperty("--document-alignment", settings.alignment);
  paper.style.setProperty(
    "--document-title-size",
    `${settings.titleFontSizePt}px`,
  );
  paper.style.setProperty(
    "--document-heading-size",
    `${settings.headingFontSizePt}px`,
  );
  paper.style.setProperty(
    "--document-subheading-size",
    `${settings.subheadingFontSizePt}px`,
  );
  paper.style.setProperty(
    "--document-table-size",
    `${settings.tableFontSizePt}px`,
  );
}

async function beginDocument(message: Record<string, unknown>): Promise<void> {
  if (
    typeof message.id !== "string" ||
    !isStoredDocument(message.meta) ||
    message.meta.id !== message.id ||
    !isOutlineSections(message.sections) ||
    message.sections.length !== message.meta.sectionCount
  ) {
    throw new Error("The extension sent invalid document metadata.");
  }
  if (bridgeNotice) {
    bridgeNotice.hidden = true;
  }
  revokeAssetObjectUrls();
  currentDocument = {
    ...message.meta,
    updatedAt: new Date().toISOString(),
  };
  applyDocumentSettings(resolveDocxSettings(currentDocument.settings));
  expectedSections = message.meta.sectionCount;
  expectedSectionIds = new Set(message.sections.map((section) => section.id));
  expectedAssetIds = new Set(message.meta.assets.map((asset) => asset.id));
  receivedSections = new Set();
  receivedAssets = new Map();
  assetErrors = new Map();
  currentAssets = new Map();
  restoredSections = new Map();
  dirtySections = new Set();
  if (editorRoot) {
    editorRoot.replaceChildren();
  }
  if (outline) {
    outline.replaceChildren();
  }
  renderOutline(message.sections);

  const previous = await loadEditorDocument(currentDocument.id);
  if (
    previous.document &&
    previous.document.updatedAt > currentDocument.extractedAt
  ) {
    if (window.confirm(`Restore saved edits for "${currentDocument.title}"?`)) {
      currentDocument = previous.document;
      currentAssets = new Map(
        previous.assets.map((asset) => [asset.id, asset]),
      );
      initializeAssetObjectUrls(currentAssets);
      for (const section of previous.sections) {
        if (!expectedSectionIds.has(section.id)) {
          continue;
        }
        restoredSections.set(section.id, section);
        receivedSections.add(section.id);
        renderSection(
          section.id,
          section.title,
          section.words,
          sanitizeHtml(section.html),
        );
      }
      setStatus("Restored saved edits.");
    } else {
      await deleteEditorDocument(currentDocument.id);
    }
  } else if (previous.document) {
    await deleteEditorDocument(currentDocument.id);
  }

  applyDocumentSettings(resolveDocxSettings(currentDocument.settings));
  await saveEditorDocument(currentDocument, [], [...currentAssets.values()]);
  document.title = `${currentDocument.title} — Page2Word`;
  if (warningBanner) {
    const warnings = currentDocument.warnings;
    warningBanner.hidden = warnings.length === 0;
    warningBanner.textContent = warnings.join(" ");
  }
  if (exportButton) {
    exportButton.disabled = true;
  }
  setStatus(
    `Loading ${expectedSections} section${expectedSections === 1 ? "" : "s"}…`,
  );
}

async function receiveSection(message: Record<string, unknown>): Promise<void> {
  if (
    !currentDocument ||
    message.id !== currentDocument.id ||
    !isRecord(message.section)
  ) {
    throw new Error("The extension sent a section for the wrong document.");
  }
  const section = message.section;
  if (
    typeof section.id !== "string" ||
    typeof section.title !== "string" ||
    typeof section.html !== "string" ||
    !Number.isInteger(section.words)
  ) {
    throw new Error("The extension sent an invalid document section.");
  }
  if (!expectedSectionIds.has(section.id)) {
    throw new Error("The extension sent an unexpected document section.");
  }
  if (receivedSections.has(section.id)) {
    return;
  }

  const cleanHtml = sanitizeHtml(section.html);
  renderSection(section.id, section.title, section.words as number, cleanHtml);
  receivedSections.add(section.id);
  setStatus(`Loaded ${receivedSections.size} of ${expectedSections} sections…`);
}

async function receiveAsset(message: Record<string, unknown>): Promise<void> {
  if (
    !currentDocument ||
    message.id !== currentDocument.id ||
    !isRecord(message.asset)
  ) {
    throw new Error("The extension sent an image for the wrong document.");
  }
  const asset = message.asset;
  if (
    typeof asset.id !== "string" ||
    !expectedAssetIds.has(asset.id) ||
    typeof asset.mime !== "string" ||
    !isImageMime(asset.mime) ||
    !Number.isInteger(asset.width) ||
    !Number.isInteger(asset.height) ||
    typeof asset.base64 !== "string"
  ) {
    throw new Error("The extension sent an invalid image asset.");
  }
  const expected = currentDocument.assets.find(
    (candidate) => candidate.id === asset.id,
  );
  if (
    !expected ||
    expected.mime !== asset.mime ||
    expected.width !== asset.width ||
    expected.height !== asset.height ||
    expected.bytes > 8 * 1024 * 1024 ||
    asset.base64.length !== Math.ceil(expected.bytes / 3) * 4
  ) {
    throw new Error("The extension image metadata did not match its payload.");
  }
  if (receivedAssets.has(asset.id)) {
    return;
  }

  const bytes = decodeBase64(asset.base64);
  if (bytes.byteLength !== expected.bytes) {
    throw new Error("The extension image payload was incomplete.");
  }
  const blob = new Blob([bytes], { type: asset.mime });
  const record: EditorAssetRecord = {
    ...expected,
    documentId: currentDocument.id,
    blob,
  };
  receivedAssets.set(asset.id, record);
  currentAssets.set(asset.id, record);
  setAssetObjectUrl(record);
  refreshImageControls();
  setStatus(
    `Loaded ${receivedAssets.size} of ${expectedAssetIds.size} images…`,
  );
}

function receiveAssetError(message: Record<string, unknown>): void {
  if (
    !currentDocument ||
    message.id !== currentDocument.id ||
    typeof message.assetId !== "string" ||
    !expectedAssetIds.has(message.assetId) ||
    typeof message.error !== "string"
  ) {
    throw new Error("The extension reported an invalid image error.");
  }
  assetErrors.set(message.assetId, message.error);
}

async function finishDocument(message: Record<string, unknown>): Promise<void> {
  if (!currentDocument || message.id !== currentDocument.id) {
    throw new Error("The extension finished the wrong document.");
  }
  const counts = message.counts;
  if (
    !isRecord(counts) ||
    counts.sections !== expectedSections ||
    receivedSections.size !== expectedSections
  ) {
    setStatus(
      `Only ${receivedSections.size} of ${expectedSections} sections arrived. Extract again to retry.`,
    );
    return;
  }
  const missingAssetIds = [...expectedAssetIds].filter(
    (id) => !receivedAssets.has(id),
  );
  if (missingAssetIds.length) {
    const missingIds = new Set(missingAssetIds);
    for (const image of editorRoot?.querySelectorAll<HTMLImageElement>(
      "img[data-p2w-asset]",
    ) ?? []) {
      if (!missingIds.has(image.dataset.p2wAsset ?? "")) {
        continue;
      }
      replaceImageWithPlaceholder(
        image,
        `[Image unavailable: ${image.alt || "image"}]`,
      );
    }
    currentDocument.assets = currentDocument.assets.filter((asset) =>
      receivedAssets.has(asset.id),
    );
    const reason = missingAssetIds
      .map((id) => assetErrors.get(id))
      .find((error): error is string => Boolean(error));
    currentDocument.warnings.push(
      `${missingAssetIds.length} image${missingAssetIds.length === 1 ? "" : "s"} did not arrive in the editor and will be exported as placeholders.${reason ? ` ${reason}` : ""}`,
    );
    if (warningBanner) {
      warningBanner.hidden = false;
      warningBanner.textContent = currentDocument.warnings.join(" ");
    }
  }
  currentDocument.updatedAt = new Date().toISOString();
  prepareInternalReferences();
  const updatedSections = [
    ...(editorRoot?.querySelectorAll<HTMLElement>("[data-p2w-section]") ?? []),
  ]
    .map((element) => sectionRecord(element))
    .filter((section): section is EditorSectionRecord => section !== null);
  await saveEditorDocument(currentDocument, updatedSections, [
    ...receivedAssets.values(),
  ]);
  currentAssets = new Map(receivedAssets);
  dirtySections.clear();
  if (saveHtmlButton) {
    saveHtmlButton.disabled = false;
  }
  if (exportButton) {
    exportButton.disabled = false;
  }
  refreshImageControls();
  setStatus("Saved locally.");
  if (extensionRuntimeAvailable) {
    void chrome.runtime
      .sendMessage({ type: "ACK", id: currentDocument.id })
      .catch((error: unknown) =>
        console.error("Could not acknowledge the editor handoff.", error),
      );
  } else {
    window.postMessage(
      { source: "page2word-editor", type: "ACK", id: currentDocument.id },
      location.origin,
    );
  }
}

async function restoreSavedDocument(id: string): Promise<void> {
  if (!/^[a-z0-9-]+$/i.test(id)) {
    throw new Error("The editor document ID is invalid.");
  }
  const saved = await loadEditorDocument(id);
  if (!saved.document) {
    if (bridgeNotice) {
      bridgeNotice.hidden = false;
    }
    return;
  }

  revokeAssetObjectUrls();
  currentDocument = saved.document;
  applyDocumentSettings(resolveDocxSettings(currentDocument.settings));
  currentAssets = new Map(saved.assets.map((asset) => [asset.id, asset]));
  if (currentAssets.size !== currentDocument.assets.length) {
    throw new Error(
      "A saved image is missing from local storage. Re-extract the article to restore it.",
    );
  }
  initializeAssetObjectUrls(currentAssets);
  expectedSections = currentDocument.sectionCount;
  expectedSectionIds = new Set(
    currentDocument.sections.map((section) => section.id),
  );
  receivedSections = new Set();
  restoredSections = new Map(
    saved.sections.map((section) => [section.id, section]),
  );
  if (editorRoot) {
    editorRoot.replaceChildren();
  }
  if (outline) {
    outline.replaceChildren();
  }
  renderOutline(currentDocument.sections);
  for (const section of saved.sections) {
    if (!expectedSectionIds.has(section.id)) {
      continue;
    }
    receivedSections.add(section.id);
    renderSection(
      section.id,
      section.title,
      section.words,
      sanitizeHtml(section.html),
    );
  }
  if (receivedSections.size !== expectedSections) {
    throw new Error(
      "The saved document is incomplete and could not be restored.",
    );
  }
  prepareInternalReferences();

  document.title = `${currentDocument.title} — Page2Word`;
  if (warningBanner) {
    warningBanner.hidden = currentDocument.warnings.length === 0;
    warningBanner.textContent = currentDocument.warnings.join(" ");
  }
  if (bridgeNotice) {
    bridgeNotice.hidden = true;
  }
  if (saveHtmlButton) {
    saveHtmlButton.disabled = false;
  }
  if (exportButton) {
    exportButton.disabled = false;
  }
  refreshImageControls();
  setStatus("Restored your saved local copy.");
}

function renderSection(
  id: string,
  title: string,
  words: number,
  html: string,
): void {
  if (!editorRoot) {
    return;
  }
  const section = document.createElement("section");
  section.id = id;
  section.dataset.p2wSection = id;
  section.contentEditable = "true";
  section.spellcheck = true;
  section.setAttribute("data-gramm", "true");
  section.setAttribute("aria-label", title);
  const saved = restoredSections.get(id);
  section.hidden = saved?.hidden ?? false;
  const content = document.createElement("template");
  content.innerHTML = html.replace(
    /(\s)src\s*=\s*(?:"p2w-asset:[^"]*"|'p2w-asset:[^']*'|p2w-asset:[^\s>]+)/gi,
    "",
  );
  normalizeArticleMarkup(content.content, document);
  applyOriginalLinkSetting(
    content.content,
    resolveDocxSettings(currentDocument?.settings).keepOriginalLinks,
  );
  addImageControls(content.content);
  renderDocumentTitle();
  section.append(content.content);
  editorRoot.append(section);
  applyAssetObjectUrls(section);
  const doneCheckbox = outline?.querySelector<HTMLInputElement>(
    `input[data-done-section="${CSS.escape(id)}"]`,
  );
  if (doneCheckbox) {
    doneCheckbox.checked = saved?.done ?? false;
  }
  updateSectionVisibilityButton(id, section.hidden);
  if (words > 0) {
    section.dataset.words = String(words);
  }
}

function renderDocumentTitle(): void {
  if (
    !editorRoot ||
    !currentDocument?.title ||
    editorRoot.querySelector("[data-p2w-document-title]")
  ) {
    return;
  }
  const title = document.createElement("h1");
  title.dataset.p2wDocumentTitle = "";
  title.textContent = currentDocument.title;
  editorRoot.append(title);
}

function renderOutline(sections: unknown): void {
  if (!outline || !Array.isArray(sections)) {
    return;
  }
  for (const candidate of sections) {
    if (
      !isRecord(candidate) ||
      typeof candidate.id !== "string" ||
      typeof candidate.title !== "string" ||
      !Number.isInteger(candidate.words)
    ) {
      continue;
    }
    const item = document.createElement("li");
    item.dataset.outlineSection = candidate.id;
    const saved = restoredSections.get(candidate.id);
    item.classList.toggle("is-hidden", saved?.hidden ?? false);
    const link = document.createElement("a");
    link.href = `#${candidate.id}`;
    link.textContent = candidate.title;
    const words = document.createElement("span");
    words.textContent = `${candidate.words} words`;
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.doneSection = candidate.id;
    label.append(checkbox, document.createTextNode(" Done"));
    const visibilityButton = document.createElement("button");
    visibilityButton.type = "button";
    visibilityButton.className = "section-visibility";
    visibilityButton.dataset.sectionVisibility = candidate.id;
    visibilityButton.textContent = saved?.hidden
      ? "Show section"
      : "Hide section";
    item.append(link, words, label, visibilityButton);
    outline.append(item);
  }
}

function prepareInternalReferences(): void {
  if (!editorRoot) {
    return;
  }
  const targets = new Map<string, string>();
  const figures = [...editorRoot.querySelectorAll<HTMLElement>("figure")];
  figures.forEach((figure, index) => {
    const number =
      referenceNumber(
        figure.querySelector("figcaption")?.textContent,
        "figure",
      ) ?? String(index + 1);
    const id = `p2w_figure_${index + 1}`;
    figure.id = id;
    if (!targets.has(`figure:${number}`)) {
      targets.set(`figure:${number}`, id);
    }
  });

  const tables = [...editorRoot.querySelectorAll<HTMLTableElement>("table")];
  tables.forEach((table, index) => {
    const number =
      referenceNumber(
        table.querySelector(":scope > caption")?.textContent,
        "table",
      ) ?? String(index + 1);
    const id = `p2w_table_${index + 1}`;
    table.id = id;
    if (!targets.has(`table:${number}`)) {
      targets.set(`table:${number}`, id);
    }
  });

  const standaloneImages = [
    ...(editorRoot.querySelectorAll<HTMLImageElement>("img") ?? []),
  ].filter((image) => !image.closest("figure"));
  standaloneImages.forEach((image, index) => {
    image.id = `p2w_image_${index + 1}`;
  });

  for (const link of editorRoot.querySelectorAll<HTMLAnchorElement>(
    'a[href^="#p2w_figure_"],a[href^="#p2w_table_"],a[href^="#p2w_image_"]',
  )) {
    link.replaceWith(...link.childNodes);
  }
  linkDocumentReferences(editorRoot, targets);
  for (const section of editorRoot.querySelectorAll<HTMLElement>(
    "[data-p2w-section]",
  )) {
    updateOutlineReferences(section);
  }
}

function linkDocumentReferences(
  root: HTMLElement,
  targets: ReadonlyMap<string, string>,
): void {
  const walker = document.createTreeWalker(
    root,
    document.defaultView?.NodeFilter.SHOW_TEXT ?? 4,
  );
  const textNodes: Text[] = [];
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    const parent = node.parentElement;
    if (
      parent &&
      !parent.closest("a,figcaption,table caption,button,[data-p2w-ui='image']")
    ) {
      textNodes.push(node);
    }
  }
  const referencePattern = /\b(Fig(?:ure)?\.?\s*\d+|Table\s+\d+)\b/gi;
  for (const node of textNodes) {
    const text = node.nodeValue ?? "";
    referencePattern.lastIndex = 0;
    let match: RegExpExecArray | null;
    let offset = 0;
    const fragment = document.createDocumentFragment();
    let changed = false;
    while ((match = referencePattern.exec(text))) {
      const target = referenceTarget(match[0]);
      const id = target ? targets.get(target.key) : undefined;
      if (!id) {
        continue;
      }
      fragment.append(document.createTextNode(text.slice(offset, match.index)));
      const link = document.createElement("a");
      link.href = `#${id}`;
      link.textContent = match[0];
      fragment.append(link);
      offset = match.index + match[0].length;
      changed = true;
    }
    if (changed) {
      fragment.append(document.createTextNode(text.slice(offset)));
      node.replaceWith(fragment);
    }
  }
}

function referenceTarget(value: string): { key: string } | undefined {
  const figure = value.match(/^Fig(?:ure)?\.?\s*(\d+)$/i);
  if (figure) {
    return { key: `figure:${figure[1]}` };
  }
  const table = value.match(/^Table\s+(\d+)$/i);
  return table ? { key: `table:${table[1]}` } : undefined;
}

function referenceNumber(
  caption: string | null | undefined,
  type: "figure" | "table",
): string | undefined {
  const pattern =
    type === "figure"
      ? /^\s*Fig(?:ure)?\.?\s*(\d{1,6})\b/i
      : /^\s*Table\s+(\d{1,6})\b/i;
  return caption?.match(pattern)?.[1];
}

function updateOutlineReferences(section: HTMLElement): void {
  const sectionId = section.dataset.p2wSection;
  if (!outline || !sectionId) {
    return;
  }
  const item = outline.querySelector<HTMLElement>(
    `[data-outline-section="${CSS.escape(sectionId)}"]`,
  );
  if (!item) {
    return;
  }
  item.querySelector(".outline-references")?.remove();

  const references: { label: string; id: string }[] = [];
  for (const figure of section.querySelectorAll<HTMLElement>("figure[id]")) {
    const number =
      referenceNumber(
        figure.querySelector("figcaption")?.textContent,
        "figure",
      ) ?? figure.id.match(/^p2w_figure_(\d+)/)?.[1];
    if (number) {
      references.push({ label: `Fig. ${number}`, id: figure.id });
    }
  }
  for (const image of section.querySelectorAll<HTMLImageElement>(
    "img[id^='p2w_image_']",
  )) {
    const number = image.id.match(/^p2w_image_(\d+)/)?.[1];
    if (number) {
      references.push({ label: `Image ${number}`, id: image.id });
    }
  }
  for (const table of section.querySelectorAll<HTMLTableElement>("table[id]")) {
    const number =
      referenceNumber(
        table.querySelector(":scope > caption")?.textContent,
        "table",
      ) ?? table.id.match(/^p2w_table_(\d+)/)?.[1];
    if (number) {
      references.push({ label: `Table ${number}`, id: table.id });
    }
  }
  if (!references.length) {
    return;
  }

  const list = document.createElement("ul");
  list.className = "outline-references";
  for (const reference of references) {
    const entry = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#${reference.id}`;
    link.textContent = reference.label;
    entry.append(link);
    list.append(entry);
  }
  item.append(list);
}

function applyOriginalLinkSetting(
  root: ParentNode,
  keepOriginalLinks: boolean,
): void {
  if (keepOriginalLinks) {
    for (const span of root.querySelectorAll<HTMLElement>(
      "span[data-p2w-original-href]",
    )) {
      if (
        span.querySelector(
          'a[href^="#p2w_figure_"],a[href^="#p2w_table_"],a[href^="#p2w_image_"]',
        )
      ) {
        continue;
      }
      const href = span.dataset.p2wOriginalHref ?? "";
      if (!isOriginalLinkHref(href)) {
        span.removeAttribute("data-p2w-original-href");
        continue;
      }
      const link = document.createElement("a");
      link.href = href;
      link.replaceChildren(...span.childNodes);
      span.replaceWith(link);
    }
    return;
  }

  for (const link of root.querySelectorAll<HTMLAnchorElement>("a[href]")) {
    const href = link.getAttribute("href") ?? "";
    if (
      !isOriginalLinkHref(href) ||
      /^#p2w_(?:figure|table|image)_/i.test(href)
    ) {
      continue;
    }
    const placeholder = document.createElement("span");
    placeholder.dataset.p2wOriginalHref = href;
    placeholder.replaceChildren(...link.childNodes);
    link.replaceWith(placeholder);
  }
}

function isOriginalLinkHref(href: string): boolean {
  return /^(?:https?:|mailto:|tel:|#[\w-]+$)/i.test(href);
}

function updateSectionVisibilityButton(id: string, hidden: boolean): void {
  const button = outline?.querySelector<HTMLButtonElement>(
    `button[data-section-visibility="${CSS.escape(id)}"]`,
  );
  if (button) {
    button.textContent = hidden ? "Show section" : "Hide section";
    button.closest("li")?.classList.toggle("is-hidden", hidden);
  }
}

function markDirty(id: string): void {
  if (!id) {
    return;
  }
  dirtySections.add(id);
  setStatus("Unsaved changes");
  if (saveTimer !== undefined) {
    window.clearTimeout(saveTimer);
  }
  saveTimer = window.setTimeout(() => void saveDirtySections(), 3_000);
}

async function saveDirtySections(): Promise<void> {
  if (!currentDocument || dirtySections.size === 0) {
    return;
  }
  if (saveTimer !== undefined) {
    window.clearTimeout(saveTimer);
    saveTimer = undefined;
  }
  const ids = [...dirtySections];
  const sections = ids
    .map((id) => {
      const element = findSection(id);
      return element ? sectionRecord(element) : null;
    })
    .filter((section): section is EditorSectionRecord => section !== null);
  if (sections.length === 0) {
    return;
  }

  setStatus("Saving…");
  try {
    await saveEditorSections(currentDocument.id, sections);
    for (const section of sections) {
      const element = findSection(section.id);
      const checkbox = outline?.querySelector<HTMLInputElement>(
        `input[data-done-section="${CSS.escape(section.id)}"]`,
      );
      const visibilityButton = outline?.querySelector<HTMLButtonElement>(
        `button[data-section-visibility="${CSS.escape(section.id)}"]`,
      );
      if (
        (element && sanitizeHtml(persistentSectionHtml(element))) ===
          section.html &&
        (checkbox?.checked ?? false) === section.done &&
        (element?.hidden ?? false) === (section.hidden ?? false) &&
        Boolean(visibilityButton?.textContent === "Show section") ===
          (section.hidden ?? false)
      ) {
        dirtySections.delete(section.id);
      }
    }
    currentDocument.updatedAt = new Date().toISOString();
    setStatus(
      dirtySections.size === 0
        ? "Saved locally."
        : "More edits are still saving…",
    );
  } catch (error) {
    showError(error);
  }
}

function sectionRecord(element: HTMLElement): EditorSectionRecord | null {
  const id = element.dataset.p2wSection;
  if (!currentDocument || !id) {
    return null;
  }
  const checkbox = outline?.querySelector<HTMLInputElement>(
    `input[data-done-section="${CSS.escape(id)}"]`,
  );
  const html = sanitizeHtml(persistentSectionHtml(element));
  return {
    ...toEditorSection(
      currentDocument.id,
      {
        id,
        title: element.getAttribute("aria-label") ?? id,
        words: Number(element.dataset.words ?? 0),
        html,
      } satisfies ContentSection,
      html,
      checkbox?.checked ?? false,
    ),
    hidden: element.hidden,
  };
}

function persistentSectionHtml(element: HTMLElement): string {
  const copy = element.cloneNode(true) as HTMLElement;
  removeEditorUi(copy);
  for (const image of copy.querySelectorAll<HTMLImageElement>("img")) {
    const assetId = image.dataset.p2wAsset;
    if (!assetId || !/^[a-z0-9-]+$/i.test(assetId)) {
      continue;
    }
    image.setAttribute("src", `p2w-asset:${assetId}`);
  }
  return copy.innerHTML;
}

function initializeAssetObjectUrls(
  assets: ReadonlyMap<string, EditorAssetRecord>,
): void {
  for (const asset of assets.values()) {
    setAssetObjectUrl(asset);
  }
}

function setAssetObjectUrl(asset: EditorAssetRecord): void {
  const previous = assetObjectUrls.get(asset.id);
  if (previous) {
    URL.revokeObjectURL(previous);
  }
  assetObjectUrls.set(asset.id, URL.createObjectURL(asset.blob));
  if (!editorRoot) {
    return;
  }
  for (const image of editorRoot.querySelectorAll<HTMLImageElement>(
    `img[data-p2w-asset="${CSS.escape(asset.id)}"]`,
  )) {
    image.src = assetObjectUrls.get(asset.id) ?? "";
  }
  refreshImageControls();
}

function addImageControls(container: ParentNode): void {
  for (const image of container.querySelectorAll<HTMLImageElement>("img")) {
    if (image.closest('[data-p2w-ui="image"]')) {
      continue;
    }
    const original =
      image.parentElement?.tagName === "A" ? image.parentElement : image;
    const assetId = image.dataset.p2wAsset ?? "";
    const imageKey = crypto.randomUUID();
    const wrapper = document.createElement("span");
    wrapper.className = "image-edit-tools";
    wrapper.contentEditable = "false";
    wrapper.dataset.p2wUi = "image";
    wrapper.dataset.assetId = assetId;
    wrapper.dataset.imageKey = imageKey;
    const content = document.createElement("span");
    content.dataset.p2wImageContent = "";
    original.replaceWith(wrapper);
    content.append(original);
    wrapper.append(content);

    const actions = document.createElement("span");
    actions.className = "image-edit-actions";
    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.dataset.imageAction = "download";
    downloadButton.dataset.assetId = assetId;
    downloadButton.dataset.imageKey = imageKey;
    downloadButton.dataset.imageName = image.alt || "article-image";
    downloadButton.textContent = "Download image";
    downloadButton.disabled =
      !(assetId && currentAssets.has(assetId)) &&
      !Boolean(image.currentSrc || image.getAttribute("src"));
    const replaceButton = document.createElement("button");
    replaceButton.type = "button";
    replaceButton.dataset.imageAction = "replace";
    replaceButton.dataset.assetId = assetId;
    replaceButton.dataset.imageKey = imageKey;
    replaceButton.textContent = "Replace image";
    replaceButton.disabled = exportButton?.disabled ?? false;
    actions.append(downloadButton, replaceButton);
    wrapper.append(actions);
  }
}

function refreshImageControls(): void {
  for (const wrapper of editorRoot?.querySelectorAll<HTMLElement>(
    '[data-p2w-ui="image"]',
  ) ?? []) {
    const assetId = wrapper.dataset.assetId ?? "";
    for (const button of wrapper.querySelectorAll<HTMLButtonElement>(
      'button[data-image-action="download"]',
    )) {
      button.dataset.assetId = assetId;
      const image = wrapper.querySelector<HTMLImageElement>("img");
      button.disabled =
        !(assetId && currentAssets.has(assetId)) &&
        !Boolean(image?.currentSrc || image?.getAttribute("src"));
    }
    for (const button of wrapper.querySelectorAll<HTMLButtonElement>(
      'button[data-image-action="replace"]',
    )) {
      button.dataset.assetId = assetId;
      button.disabled = exportButton?.disabled ?? false;
    }
  }
}

async function downloadArticleImage(
  assetId: string | undefined,
  imageKey: string,
  name: string,
): Promise<void> {
  const asset = assetId ? currentAssets.get(assetId) : undefined;
  if (asset) {
    downloadBlob(
      asset.blob,
      `${safeFilename(name)}.${imageExtension(asset.mime)}`,
    );
    setStatus("Image downloaded.");
    return;
  }

  const image = editorRoot?.querySelector<HTMLImageElement>(
    `[data-p2w-ui="image"][data-image-key="${CSS.escape(imageKey)}"] img`,
  );
  const source = image?.currentSrc || image?.getAttribute("src");
  if (!source || source.startsWith("p2w-asset:")) {
    throw new Error("The image is not available to download.");
  }
  let response: Response;
  try {
    response = await fetch(source);
  } catch {
    throw new Error(
      "The publisher blocked this image download. Try replacing the image with a local copy.",
    );
  }
  if (!response.ok) {
    throw new Error(`The image download failed (${response.status}).`);
  }
  const blob = await response.blob();
  if (!blob.type.startsWith("image/")) {
    throw new Error("The image source returned a non-image file.");
  }
  if (!blob.size || blob.size > 8 * 1024 * 1024) {
    throw new Error("The image download is empty or larger than 8 MB.");
  }
  downloadBlob(
    blob,
    `${safeFilename(name)}.${imageExtensionForMime(blob.type)}`,
  );
  setStatus("Image downloaded.");
}

async function replaceImageAsset(
  assetId: string | undefined,
  imageKey: string,
  file: File,
): Promise<void> {
  if (!currentDocument || !editorRoot) {
    throw new Error("Open an article before replacing an image.");
  }
  if (!isImageMime(file.type)) {
    throw new Error("Choose a PNG, JPEG, or GIF image.");
  }
  if (file.size < 1 || file.size > 8 * 1024 * 1024) {
    throw new Error("The replacement image must be smaller than 8 MB.");
  }

  const image = editorRoot.querySelector<HTMLImageElement>(
    `[data-p2w-ui="image"][data-image-key="${CSS.escape(imageKey)}"] img`,
  );
  if (!image) {
    throw new Error("The image to replace is no longer in this article.");
  }
  const dimensions = await readImageDimensions(file);
  const id = assetId || `asset-${crypto.randomUUID()}`;
  image.dataset.p2wAsset = id;
  const wrapper = image.closest<HTMLElement>('[data-p2w-ui="image"]');
  if (wrapper) {
    wrapper.dataset.assetId = id;
  }

  const record: EditorAssetRecord = {
    id,
    documentId: currentDocument.id,
    mime: file.type,
    width: dimensions.width,
    height: dimensions.height,
    bytes: file.size,
    blob: file,
  };
  currentAssets.set(id, record);
  const assetMetadata: ExtractedAsset = {
    id,
    mime: record.mime,
    width: record.width,
    height: record.height,
    bytes: record.bytes,
  };
  const metadataIndex = currentDocument.assets.findIndex(
    (asset) => asset.id === id,
  );
  if (metadataIndex < 0) {
    currentDocument.assets.push(assetMetadata);
  } else {
    currentDocument.assets[metadataIndex] = assetMetadata;
  }
  setAssetObjectUrl(record);
  refreshImageControls();
  const section = findEditorSection(image);
  if (section && !assetId) {
    markDirty(section.dataset.p2wSection ?? "");
  }
  await saveEditorDocument(currentDocument, [], [...currentAssets.values()]);
  if (section && !assetId) {
    await saveDirtySections();
    if (dirtySections.has(section.dataset.p2wSection ?? "")) {
      throw new Error("The replacement image reference could not be saved.");
    }
  }
  setStatus("Image replaced and saved locally.");
}

function readImageDimensions(
  file: File,
): Promise<{ width: number; height: number }> {
  const url = URL.createObjectURL(file);
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      if (!image.naturalWidth || !image.naturalHeight) {
        reject(new Error("The selected file is not a valid image."));
        return;
      }
      resolve({ width: image.naturalWidth, height: image.naturalHeight });
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("The selected file could not be decoded as an image."));
    };
    image.src = url;
  });
}

function removeEditorUi(root: HTMLElement): void {
  for (const wrapper of root.querySelectorAll<HTMLElement>(
    '[data-p2w-ui="image"]',
  )) {
    const content = wrapper.querySelector<HTMLElement>(
      "[data-p2w-image-content]",
    );
    if (!content) {
      wrapper.remove();
      continue;
    }
    content.querySelectorAll("[data-p2w-image-key]").forEach((image) => {
      image.removeAttribute("data-p2w-image-key");
    });
    wrapper.replaceWith(...content.childNodes);
  }
}

function replaceImageWithPlaceholder(
  image: HTMLImageElement,
  text: string,
): void {
  const placeholder = document.createElement("span");
  placeholder.textContent = text;
  const wrapper = image.closest<HTMLElement>('[data-p2w-ui="image"]');
  if (wrapper) {
    wrapper.replaceWith(placeholder);
  } else {
    image.replaceWith(placeholder);
  }
}

function imageExtension(mime: ExtractedAsset["mime"]): string {
  return imageExtensionForMime(mime);
}

function imageExtensionForMime(mime: string): string {
  switch (mime) {
    case "image/jpeg":
      return "jpg";
    case "image/gif":
      return "gif";
    case "image/webp":
      return "webp";
    case "image/svg+xml":
      return "svg";
    default:
      return "png";
  }
}

function applyAssetObjectUrls(container: ParentNode): void {
  for (const image of container.querySelectorAll<HTMLImageElement>(
    "img[data-p2w-asset]",
  )) {
    const objectUrl = assetObjectUrls.get(image.dataset.p2wAsset ?? "");
    if (objectUrl) {
      image.src = objectUrl;
    } else if (image.getAttribute("src")?.startsWith("p2w-asset:")) {
      image.removeAttribute("src");
    }
  }
}

function revokeAssetObjectUrls(): void {
  for (const url of assetObjectUrls.values()) {
    URL.revokeObjectURL(url);
  }
  assetObjectUrls.clear();
}

async function loadHtmlFile(file: File): Promise<void> {
  try {
    await loadHtmlContent(
      await file.text(),
      file.name.replace(/\.html?$/i, "") || "Imported document",
    );
  } catch (error) {
    showError(error);
  }
}

async function loadHtmlContent(
  rawHtml: string,
  fallbackTitle: string,
): Promise<void> {
  try {
    const source = new DOMParser().parseFromString(rawHtml, "text/html");
    const html = sanitizeHtml(source.body.innerHTML);
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const id = `local-${crypto.randomUUID()}`;
    const title = source.title.trim() || fallbackTitle;
    const text = parsed.body.textContent ?? "";
    const section: ContentSection = {
      id: "section-1",
      title,
      words: text.trim().split(/\s+/).filter(Boolean).length,
      html,
    };
    const importedDocument: EditorDocumentRecord = {
      id,
      title,
      sourceUrl: "",
      extractedAt: new Date().toISOString(),
      lang: source.documentElement.lang || "en",
      sectionCount: 1,
      sections: [
        { id: section.id, title: section.title, words: section.words },
      ],
      assets: [],
      stats: {
        words: section.words,
        images: parsed.body.querySelectorAll("img").length,
        tables: parsed.body.querySelectorAll("table").length,
        charts: parsed.body.querySelectorAll("figure").length,
      },
      warnings: [],
      settings: {},
      updatedAt: new Date().toISOString(),
    };
    revokeAssetObjectUrls();
    currentAssets = new Map();
    receivedAssets = new Map();
    expectedAssetIds = new Set();
    currentDocument = importedDocument;
    restoredSections = new Map();
    dirtySections = new Set();
    applyDocumentSettings(resolveDocxSettings(currentDocument.settings));
    history.replaceState(
      null,
      "",
      `${location.pathname}${location.search}#doc=${encodeURIComponent(id)}`,
    );
    expectedSections = 1;
    editorRoot?.replaceChildren();
    outline?.replaceChildren();
    renderOutline(importedDocument.sections);
    renderSection(section.id, section.title, section.words, html);
    prepareInternalReferences();
    const savedSection = editorRoot?.querySelector<HTMLElement>(
      `[data-p2w-section="${CSS.escape(section.id)}"]`,
    );
    const sectionRecordToSave = savedSection
      ? sectionRecord(savedSection)
      : null;
    if (!sectionRecordToSave) {
      throw new Error("The imported document section could not be saved.");
    }
    await saveEditorDocument(importedDocument, [sectionRecordToSave]);
    expectedSectionIds = new Set([section.id]);
    receivedSections = new Set([section.id]);
    if (bridgeNotice) {
      bridgeNotice.hidden = true;
    }
    if (warningBanner) {
      warningBanner.hidden = true;
      warningBanner.textContent = "";
    }
    if (saveHtmlButton) {
      saveHtmlButton.disabled = false;
    }
    if (exportButton) {
      exportButton.disabled = false;
    }
    refreshImageControls();
    document.title = `${section.title} — Page2Word`;
    setStatus("Opened HTML file. Changes will be saved locally.");
  } catch (error) {
    showError(error);
  }
}

async function exportWordDocument(): Promise<void> {
  if (!currentDocument || !editorRoot || !exportDialog) {
    setStatus("Open or extract an article before exporting.");
    return;
  }
  const sectionElements = [
    ...editorRoot.querySelectorAll<HTMLElement>("[data-p2w-section]"),
  ].filter((section) => !section.hidden);
  if (!sectionElements.length) {
    setStatus("There is no article content to export.");
    return;
  }

  const sections = sectionElements.map((section) => {
    const copy = section.cloneNode(true) as HTMLElement;
    removeEditorUi(copy);
    return copy.innerHTML;
  });
  const settings = resolveDocxSettings(
    currentDocument.settings ?? DEFAULT_DOCX_SETTINGS,
  );
  const metadata: DocxMetadata = {
    title: currentDocument.title,
    sourceUrl: currentDocument.sourceUrl,
    byline: currentDocument.byline,
    extractedAt: currentDocument.extractedAt,
    lang: currentDocument.lang,
  };
  const controller = new AbortController();
  activeExportController = controller;
  if (exportProgress) {
    exportProgress.hidden = false;
    exportProgress.removeAttribute("value");
  }
  if (exportSummary) {
    exportSummary.hidden = true;
    exportSummary.textContent = "";
  }
  if (cancelExportButton) {
    cancelExportButton.hidden = false;
    cancelExportButton.disabled = false;
  }
  if (closeExportButton) {
    closeExportButton.hidden = true;
  }
  if (exportButton) {
    exportButton.disabled = true;
  }

  exportDialog.showModal();
  setExportStage("Preparing sanitized content…");
  setStatus("Preparing Word export…");

  try {
    const blocks = await parseHtmlToDocxBlocks(
      sections,
      document,
      updateExportProgress,
      controller.signal,
      { keepOriginalLinks: settings.keepOriginalLinks },
    );
    if (controller.signal.aborted) {
      throw new DOMException("Export cancelled.", "AbortError");
    }

    setExportStage("Preparing embedded images…");
    const imageAssets: DocxImageAsset[] = [];
    for (const asset of currentAssets.values()) {
      if (controller.signal.aborted) {
        throw new DOMException("Export cancelled.", "AbortError");
      }
      imageAssets.push({
        id: asset.id,
        mime: asset.mime,
        width: asset.width,
        height: asset.height,
        bytes: await asset.blob.arrayBuffer(),
      });
    }
    const blob = await buildDocxInWorker(
      blocks,
      metadata,
      settings,
      imageAssets,
      controller,
    );
    if (controller.signal.aborted) {
      throw new DOMException("Export cancelled.", "AbortError");
    }
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    downloadBlob(blob, `${safeFilename(metadata.title)}-${date}.docx`);

    const imageCount = sectionElements.reduce(
      (count, section) =>
        count + section.querySelectorAll("img[data-p2w-asset]").length,
      0,
    );
    const unavailableImageCount = sectionElements.reduce(
      (count, section) =>
        count + section.querySelectorAll("img:not([data-p2w-asset])").length,
      0,
    );
    const tableCount = sectionElements.reduce(
      (count, section) => count + section.querySelectorAll("table").length,
      0,
    );
    const wordCount = sections
      .join(" ")
      .replace(/<[^>]+>/g, " ")
      .trim()
      .split(/\s+/)
      .filter(Boolean).length;
    const warnings = [...currentDocument.warnings];
    if (unavailableImageCount) {
      warnings.push(
        `${unavailableImageCount} image${unavailableImageCount === 1 ? "" : "s"} could not be embedded and were replaced with placeholders.`,
      );
    }
    if (exportProgress) {
      exportProgress.hidden = true;
    }
    if (exportSummary) {
      exportSummary.hidden = false;
      exportSummary.textContent = [
        `Downloaded ${blob.size.toLocaleString()} bytes.`,
        `${wordCount.toLocaleString()} words · ${imageCount} embedded images · ${tableCount} formatted tables.`,
        ...warnings,
      ].join(" ");
    }
    setExportStage("Word document ready.");
    setStatus("Word document downloaded.");
    if (cancelExportButton) {
      cancelExportButton.hidden = true;
    }
    if (closeExportButton) {
      closeExportButton.hidden = false;
    }
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      if (exportProgress) {
        exportProgress.hidden = true;
      }
      if (exportSummary) {
        exportSummary.hidden = false;
        exportSummary.textContent =
          "Export cancelled. Your article remains saved locally.";
      }
      setExportStage("Export cancelled.");
      setStatus("Word export cancelled.");
    } else {
      showError(error);
      setExportStage("Export failed.");
      if (exportSummary) {
        exportSummary.hidden = false;
        exportSummary.textContent =
          error instanceof Error
            ? error.message
            : "The Word document could not be generated.";
      }
    }
    if (cancelExportButton) {
      cancelExportButton.hidden = true;
    }
    if (closeExportButton) {
      closeExportButton.hidden = false;
    }
  } finally {
    activeExportWorker?.terminate();
    activeExportWorker = undefined;
    activeExportController = undefined;
    if (exportButton) {
      exportButton.disabled = false;
    }
    refreshImageControls();
  }
}

function buildDocxInWorker(
  blocks: DocxBlock[],
  metadata: DocxMetadata,
  settings: DocxSettings,
  assets: DocxImageAsset[],
  controller: AbortController,
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("docx-worker.js", location.href));
    activeExportWorker = worker;

    const cleanup = (): void => {
      controller.signal.removeEventListener("abort", onAbort);
      worker.removeEventListener("message", onMessage);
      worker.removeEventListener("error", onError);
      worker.terminate();
      if (activeExportWorker === worker) {
        activeExportWorker = undefined;
      }
    };
    const onAbort = (): void => {
      cleanup();
      reject(new DOMException("Export cancelled.", "AbortError"));
    };
    const onMessage = (event: MessageEvent): void => {
      if (event.data?.type === "PROGRESS") {
        setExportStage("Building the Word document…");
        if (exportProgress) {
          exportProgress.hidden = false;
          exportProgress.removeAttribute("value");
        }
      } else if (
        event.data?.type === "DONE" &&
        event.data.blob instanceof Blob
      ) {
        cleanup();
        resolve(event.data.blob);
      } else if (event.data?.type === "ERROR") {
        cleanup();
        reject(
          new Error(
            typeof event.data.error === "string"
              ? event.data.error
              : "The Word document could not be generated.",
          ),
        );
      }
    };
    const onError = (event: ErrorEvent): void => {
      cleanup();
      reject(new Error(event.message || "The Word export worker failed."));
    };

    controller.signal.addEventListener("abort", onAbort, { once: true });
    worker.addEventListener("message", onMessage);
    worker.addEventListener("error", onError);
    worker.postMessage(
      { type: "EXPORT", blocks, metadata, settings, assets },
      assets.map((asset) => asset.bytes),
    );
  });
}

function updateExportProgress(progress: ExportProgress): void {
  if (progress.stage !== "preparing") {
    return;
  }
  setExportStage(
    progress.total
      ? `Preparing section ${progress.completed} of ${progress.total}…`
      : "Preparing sanitized content…",
  );
  if (exportProgress && progress.total > 0) {
    exportProgress.max = progress.total;
    exportProgress.value = Math.min(progress.total, progress.completed);
  }
}

function setExportStage(message: string): void {
  if (exportStage) {
    exportStage.textContent = message;
  }
}

function findSection(id: string): HTMLElement | null {
  return (
    editorRoot?.querySelector<HTMLElement>(
      `[data-p2w-section="${CSS.escape(id)}"]`,
    ) ?? null
  );
}

function findEditorSection(target: EventTarget | null): HTMLElement | null {
  if (!(target instanceof Element)) {
    return null;
  }
  const section = target.closest<HTMLElement>("[data-p2w-section]");
  return section && editorRoot?.contains(section) ? section : null;
}

function selectedEditorSection(
  selection: Selection | null,
): HTMLElement | null {
  if (!selection || selection.isCollapsed) {
    return null;
  }
  const anchor = sectionForNode(selection.anchorNode);
  const focus = sectionForNode(selection.focusNode);
  return anchor && anchor === focus ? anchor : null;
}

function sectionForNode(node: Node | null): HTMLElement | null {
  if (!node) {
    return null;
  }
  const element = node.nodeType === 1 ? (node as Element) : node.parentElement;
  const section = element?.closest<HTMLElement>("[data-p2w-section]");
  return section && editorRoot?.contains(section) ? section : null;
}

function setStatus(message: string): void {
  if (status) {
    status.textContent = message;
  }
}

function showError(error: unknown): void {
  const message =
    error instanceof Error
      ? error.message
      : "The editor could not complete this action.";
  setStatus(message);
}

function isStoredDocument(value: unknown): value is StoredDocument {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    typeof value.title === "string" &&
    typeof value.sourceUrl === "string" &&
    typeof value.extractedAt === "string" &&
    typeof value.sectionCount === "number" &&
    Number.isInteger(value.sectionCount) &&
    value.sectionCount >= 0 &&
    isOutlineSections(value.sections) &&
    Array.isArray(value.assets) &&
    value.assets.every(isStoredAsset) &&
    Array.isArray(value.warnings) &&
    isRecord(value.stats) &&
    typeof value.stats.words === "number" &&
    typeof value.stats.images === "number" &&
    typeof value.stats.tables === "number" &&
    typeof value.stats.charts === "number" &&
    isRecord(value.settings)
  );
}

function isStoredAsset(value: unknown): value is ExtractedAsset {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    /^[a-z0-9-]+$/i.test(value.id) &&
    typeof value.mime === "string" &&
    isImageMime(value.mime) &&
    typeof value.width === "number" &&
    Number.isInteger(value.width) &&
    typeof value.height === "number" &&
    Number.isInteger(value.height) &&
    typeof value.bytes === "number" &&
    Number.isInteger(value.bytes) &&
    value.width >= 0 &&
    value.height >= 0 &&
    value.width <= 100_000 &&
    value.height <= 100_000 &&
    value.bytes > 0 &&
    value.bytes <= 8 * 1024 * 1024
  );
}

function isImageMime(value: string): value is ExtractedAsset["mime"] {
  return (
    value === "image/png" || value === "image/jpeg" || value === "image/gif"
  );
}

function isOutlineSections(
  value: unknown,
): value is StoredDocument["sections"] {
  return (
    Array.isArray(value) &&
    value.every(
      (section) =>
        isRecord(section) &&
        typeof section.id === "string" &&
        typeof section.title === "string" &&
        typeof section.words === "number" &&
        Number.isInteger(section.words),
    )
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function escapeHtml(value: string): string {
  return value.replace(
    /[&<>"']/g,
    (character) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[character] ?? character,
  );
}

function safeFilename(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-").trim() || "page2word";
}

function download(content: string, filename: string, mime: string): void {
  const objectUrl = URL.createObjectURL(
    new Blob([content], { type: `${mime};charset=utf-8` }),
  );
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 1_000);
}

function downloadBlob(blob: Blob, filename: string): void {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 60_000);
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(offset, offset + chunkSize),
    );
  }
  return btoa(binary);
}

function decodeBase64(value: string): ArrayBuffer {
  const binary = atob(value);
  const buffer = new ArrayBuffer(binary.length);
  const bytes = new Uint8Array(buffer);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return buffer;
}
