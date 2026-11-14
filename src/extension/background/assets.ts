import type { ExtractedAsset, ExtractedDoc } from "../../shared/types";

const MAX_ASSET_BYTES = 8 * 1024 * 1024;
const MAX_TOTAL_BYTES = 40 * 1024 * 1024;
const MAX_ASSETS = 100;
const FETCH_TIMEOUT_MS = 15_000;

export interface CapturedAsset {
  metadata: ExtractedAsset;
  base64: string;
}

interface ImageSource {
  src: string;
  width: number;
  height: number;
  alt: string;
}

interface CapturedImage extends CapturedAsset {
  source: string;
}

export async function captureDocumentAssets(
  document: ExtractedDoc,
  fetchImage: typeof fetch = fetch,
): Promise<CapturedAsset[]> {
  const sources = collectImageSources(document.contentSections);
  const uniqueSources = [
    ...new Map(sources.map((source) => [source.src, source])).values(),
  ];
  const captured: CapturedImage[] = [];
  const failures = new Map<string, string>();
  let totalBytes = 0;

  for (let offset = 0; offset < uniqueSources.length; offset += 4) {
    const batch = uniqueSources.slice(offset, offset + 4);
    const results = await Promise.all(
      batch.map(async (source) => {
        try {
          return await fetchAndEncode(source, fetchImage);
        } catch (error) {
          return {
            source: source.src,
            error:
              error instanceof Error ? error.message : "Image capture failed.",
          };
        }
      }),
    );

    for (const result of results) {
      if ("error" in result) {
        failures.set(result.source, result.error);
        continue;
      }
      if (captured.length >= MAX_ASSETS) {
        failures.set(result.source, "The image-count limit was reached.");
        continue;
      }
      if (totalBytes + result.metadata.bytes > MAX_TOTAL_BYTES) {
        failures.set(
          result.source,
          "The document image-size limit was reached.",
        );
        continue;
      }
      totalBytes += result.metadata.bytes;
      captured.push(result);
    }
  }

  const assets = new Map(captured.map((asset) => [asset.source, asset]));
  for (const section of document.contentSections) {
    section.html = rewriteImages(section.html, assets, failures);
  }
  document.html = document.contentSections
    .map((section) => section.html)
    .join("");
  document.assets = captured.map(({ metadata }) => metadata);
  if (failures.size) {
    document.warnings.push(
      `${failures.size} image${failures.size === 1 ? "" : "s"} could not be embedded; unavailable images are labeled in the article.`,
    );
  }
  return captured.map(({ metadata, base64 }) => ({ metadata, base64 }));
}

async function fetchAndEncode(
  source: ImageSource,
  fetchImage: typeof fetch,
): Promise<CapturedImage> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
  let blob: Blob;
  try {
    const response = await fetchImage(source.src, {
      credentials: "include",
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`Image request returned HTTP ${response.status}.`);
    }
    blob = await response.blob();
  } finally {
    clearTimeout(timeout);
  }
  if (!blob.size || blob.size > MAX_ASSET_BYTES) {
    throw new Error(
      blob.size
        ? "The image exceeds the 8 MB size limit."
        : "The image is empty.",
    );
  }

  const normalized = await normalizeImage(blob, source.src);
  if (
    !normalized.bytes.byteLength ||
    normalized.bytes.byteLength > MAX_ASSET_BYTES
  ) {
    throw new Error("The converted image exceeds the 8 MB size limit.");
  }
  const id = `asset-${crypto.randomUUID()}`;
  const dimensions = readDimensions(normalized.bytes, normalized.mime);
  return {
    source: source.src,
    metadata: {
      id,
      mime: normalized.mime,
      width: dimensions.width || source.width,
      height: dimensions.height || source.height,
      bytes: normalized.bytes.byteLength,
    },
    base64: bytesToBase64(normalized.bytes),
  };
}

async function normalizeImage(
  blob: Blob,
  source: string,
): Promise<{ mime: ExtractedAsset["mime"]; bytes: Uint8Array }> {
  const declaredMime = (blob.type || mimeFromUrl(source)).toLowerCase();
  if (
    declaredMime === "image/png" ||
    declaredMime === "image/jpeg" ||
    declaredMime === "image/gif"
  ) {
    return {
      mime: declaredMime,
      bytes: new Uint8Array(await blob.arrayBuffer()),
    };
  }

  if (
    !["image/webp", "image/avif", "image/svg+xml", "image/bmp"].includes(
      declaredMime,
    ) ||
    typeof createImageBitmap !== "function" ||
    typeof OffscreenCanvas !== "function"
  ) {
    throw new Error(`Unsupported image format: ${declaredMime || "unknown"}.`);
  }

  const bitmap = await createImageBitmap(blob);
  try {
    const scale = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height));
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = new OffscreenCanvas(width, height);
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Could not prepare the image for Word export.");
    }
    context.drawImage(bitmap, 0, 0, width, height);
    const outputMime =
      declaredMime === "image/webp" || declaredMime === "image/avif"
        ? "image/jpeg"
        : "image/png";
    const converted = await canvas.convertToBlob({
      type: outputMime,
      ...(outputMime === "image/jpeg" ? { quality: 0.85 } : {}),
    });
    return {
      mime: outputMime,
      bytes: new Uint8Array(await converted.arrayBuffer()),
    };
  } finally {
    bitmap.close();
  }
}

function collectImageSources(
  sections: ExtractedDoc["contentSections"],
): ImageSource[] {
  const sources: ImageSource[] = [];
  for (const section of sections) {
    const tags = section.html.match(/<img\b[^>]*>/gi) ?? [];
    for (const tag of tags) {
      const src = decodeHtmlAttribute(readAttribute(tag, "src") ?? "");
      if (!src || !/^(?:https?:|data:image\/)/i.test(src)) {
        continue;
      }
      sources.push({
        src,
        width: positiveDimension(readAttribute(tag, "width")),
        height: positiveDimension(readAttribute(tag, "height")),
        alt: decodeHtmlAttribute(readAttribute(tag, "alt") ?? ""),
      });
    }
  }
  return sources;
}

function rewriteImages(
  html: string,
  assets: Map<string, CapturedImage>,
  failures: Map<string, string>,
): string {
  return html.replace(/<img\b[^>]*>/gi, (tag) => {
    const src = decodeHtmlAttribute(readAttribute(tag, "src") ?? "");
    const captured = assets.get(src);
    if (!captured) {
      const reason = failures.get(src);
      if (!reason) {
        return tag;
      }
      const alt = decodeHtmlAttribute(readAttribute(tag, "alt") ?? "");
      return `<span>[Image unavailable${alt ? `: ${escapeText(alt)}` : ""}]</span>`;
    }

    let updated = setAttribute(tag, "src", `p2w-asset:${captured.metadata.id}`);
    updated = setAttribute(updated, "data-p2w-asset", captured.metadata.id);
    if (captured.metadata.width) {
      updated = setAttribute(updated, "width", String(captured.metadata.width));
    }
    if (captured.metadata.height) {
      updated = setAttribute(
        updated,
        "height",
        String(captured.metadata.height),
      );
    }
    return updated;
  });
}

function readAttribute(tag: string, name: string): string | null {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tag.match(
    new RegExp(
      `(?:^|\\s)${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`,
      "i",
    ),
  );
  return match?.[1] ?? match?.[2] ?? match?.[3] ?? null;
}

function setAttribute(tag: string, name: string, value: string): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const withoutAttribute = tag.replace(
    new RegExp(`\\s${escaped}\\s*=\\s*(?:"[^"]*"|'[^']*'|[^\\s>]+)`, "i"),
    "",
  );
  return withoutAttribute.replace(
    /\s*\/?>$/,
    (ending) => ` ${name}="${escapeAttribute(value)}"${ending.trimStart()}`,
  );
}

function readDimensions(
  bytes: Uint8Array,
  mime: ExtractedAsset["mime"],
): { width: number; height: number } {
  if (mime === "image/png" && bytes.length >= 24) {
    return {
      width: new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      ).getUint32(16),
      height: new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      ).getUint32(20),
    };
  }
  if (mime === "image/gif" && bytes.length >= 10) {
    return {
      width: bytes[6] | (bytes[7] << 8),
      height: bytes[8] | (bytes[9] << 8),
    };
  }
  return { width: 0, height: 0 };
}

function mimeFromUrl(source: string): string {
  try {
    const pathname = new URL(source).pathname.toLowerCase();
    if (pathname.endsWith(".png")) return "image/png";
    if (pathname.endsWith(".jpg") || pathname.endsWith(".jpeg"))
      return "image/jpeg";
    if (pathname.endsWith(".gif")) return "image/gif";
    if (pathname.endsWith(".webp")) return "image/webp";
    if (pathname.endsWith(".avif")) return "image/avif";
    if (pathname.endsWith(".svg")) return "image/svg+xml";
    if (pathname.endsWith(".bmp")) return "image/bmp";
  } catch {
    return "";
  }
  return "";
}

function positiveDimension(value: string | null): number {
  const parsed = Number.parseInt(value ?? "", 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : 0;
}

function decodeHtmlAttribute(value: string): string {
  return value
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
}

function escapeAttribute(value: string): string {
  return value.replace(
    /[&"<>]/g,
    (character) =>
      ({
        "&": "&amp;",
        '"': "&quot;",
        "<": "&lt;",
        ">": "&gt;",
      })[character] ?? character,
  );
}

function escapeText(value: string): string {
  return value.replace(
    /[&<>]/g,
    (character) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
      })[character] ?? character,
  );
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
