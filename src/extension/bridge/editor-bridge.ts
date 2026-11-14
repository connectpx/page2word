import {
  documentAssetKey,
  documentMetaKey,
  documentSectionKey,
} from "../../shared/storage-keys";
import { DOCX_SETTINGS_KEY } from "../../shared/html-to-docx/settings";

const EDITOR_SOURCE = "page2word-editor";
const BRIDGE_SOURCE = "page2word";
const docId = new URLSearchParams(location.hash.slice(1)).get("doc");
let streaming = false;

window.addEventListener("message", (event: MessageEvent<unknown>) => {
  if (
    event.source !== window ||
    event.origin !== location.origin ||
    !isRecord(event.data)
  ) {
    return;
  }
  const message = event.data;
  if (message.source !== EDITOR_SOURCE) {
    return;
  }

  if (message.type === "READY" && !streaming) {
    streaming = true;
    void streamDocument().catch((error: unknown) =>
      post({
        type: "ERROR",
        error:
          error instanceof Error
            ? error.message
            : "The extension could not load this document.",
      }),
    );
  } else if (message.type === "ACK" && message.id === docId) {
    void chrome.runtime
      .sendMessage({ type: "ACK", id: docId })
      .catch((error: unknown) =>
        console.error("Could not acknowledge the editor handoff.", error),
      );
  } else if (message.type === "OPEN_OPTIONS") {
    void chrome.runtime
      .sendMessage({ type: "OPEN_OPTIONS" })
      .catch((error: unknown) =>
        console.error("Could not open Page2Word settings.", error),
      );
  }
});

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "local" || !changes[DOCX_SETTINGS_KEY]) {
    return;
  }
  post({ type: "SETTINGS", settings: changes[DOCX_SETTINGS_KEY].newValue });
});

async function streamDocument(): Promise<void> {
  if (!docId || !/^[a-z0-9-]+$/i.test(docId)) {
    post({
      type: "ERROR",
      error: "No extracted document was found for this editor tab.",
    });
    return;
  }

  const key = documentMetaKey(docId);
  const stored = await chrome.storage.local.get(key);
  const meta = stored[key];
  if (
    !isRecord(meta) ||
    meta.id !== docId ||
    typeof meta.sectionCount !== "number" ||
    !Number.isInteger(meta.sectionCount) ||
    !Array.isArray(meta.sections) ||
    !Array.isArray(meta.assets) ||
    meta.sectionCount < 0 ||
    meta.sectionCount > 10_000 ||
    meta.sections.length !== meta.sectionCount
  ) {
    post({
      type: "ERROR",
      error: "The extracted document has expired. Extract the page again.",
    });
    return;
  }

  if (meta.sectionsAvailable === false) {
    post({ type: "REOPEN", id: docId });
    return;
  }

  const sectionCount = meta.sectionCount as number;
  const sectionKeys = Array.from({ length: sectionCount }, (_, index) =>
    documentSectionKey(docId, index),
  );
  const payload = await chrome.storage.local.get(sectionKeys);
  post({
    type: "DOC_BEGIN",
    id: docId,
    meta,
    settings: isRecord(meta.settings) ? meta.settings : {},
    sections: meta.sections,
  });

  for (let index = 0; index < sectionKeys.length; index += 1) {
    const section = payload[sectionKeys[index]];
    if (!isRecord(section) || typeof section.html !== "string") {
      post({
        type: "ERROR",
        error: "A document section is missing. Please extract the page again.",
      });
      return;
    }
    if (
      typeof section.id !== "string" ||
      typeof section.title !== "string" ||
      typeof section.words !== "number"
    ) {
      post({
        type: "ERROR",
        error: "A document section is invalid. Please extract the page again.",
      });
      return;
    }
    post({
      type: "SECTION",
      id: docId,
      index,
      section: {
        id: section.id,
        title: section.title,
        words: section.words,
        html: section.html,
      },
    });
  }

  for (const asset of meta.assets) {
    if (
      !isRecord(asset) ||
      typeof asset.id !== "string" ||
      !/^[a-z0-9-]+$/i.test(asset.id) ||
      typeof asset.mime !== "string" ||
      !["image/png", "image/jpeg", "image/gif"].includes(asset.mime) ||
      typeof asset.width !== "number" ||
      !Number.isInteger(asset.width) ||
      asset.width < 0 ||
      asset.width > 100_000 ||
      typeof asset.height !== "number" ||
      !Number.isInteger(asset.height) ||
      asset.height < 0 ||
      asset.height > 100_000 ||
      typeof asset.bytes !== "number" ||
      !Number.isInteger(asset.bytes) ||
      asset.bytes < 1 ||
      asset.bytes > 8 * 1024 * 1024
    ) {
      post({
        type: "ERROR",
        error: "An extracted image asset is invalid. Extract the page again.",
      });
      return;
    }

    let base64: unknown;
    const assetKey = documentAssetKey(docId, asset.id);
    try {
      const storedAsset = await chrome.storage.local.get(assetKey);
      base64 = storedAsset[assetKey];
    } catch (error) {
      postAssetError(
        docId,
        asset.id,
        error instanceof Error
          ? error.message
          : "The image could not be loaded.",
      );
      continue;
    }
    if (typeof base64 !== "string") {
      postAssetError(
        docId,
        asset.id,
        "The image is missing from temporary extension storage.",
      );
      continue;
    }
    if (
      base64.length !== Math.ceil(asset.bytes / 3) * 4 ||
      !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)
    ) {
      postAssetError(
        docId,
        asset.id,
        "The image data is incomplete in temporary extension storage.",
      );
      continue;
    }
    window.postMessage(
      {
        source: BRIDGE_SOURCE,
        type: "ASSET",
        id: docId,
        asset: {
          id: asset.id,
          mime: asset.mime,
          width: asset.width,
          height: asset.height,
          base64,
        },
      },
      location.origin,
    );
  }

  post({
    type: "DOC_END",
    id: docId,
    counts: { sections: sectionCount, assets: meta.assets.length },
  });
}

function postAssetError(id: string, assetId: string, error: string): void {
  post({ type: "ASSET_ERROR", id, assetId, error });
}

function post(message: Record<string, unknown>): void {
  window.postMessage({ source: BRIDGE_SOURCE, ...message }, location.origin);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
