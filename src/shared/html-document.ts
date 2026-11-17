import type { ContentSection } from "./types";

export async function decodeHtmlFile(file: Blob): Promise<string> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const prefix = new TextDecoder("windows-1252").decode(
    bytes.subarray(0, 4096),
  );
  const declaredCharset = prefix.match(
    /<meta\b[^>]*(?:charset\s*=\s*["']?\s*([^\s"'/>;]+)|content\s*=\s*["'][^"']*charset\s*=\s*([^\s;"']+))/i,
  );
  const bomCharset =
    bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
      ? "utf-8"
      : bytes[0] === 0xff && bytes[1] === 0xfe
        ? "utf-16le"
        : bytes[0] === 0xfe && bytes[1] === 0xff
          ? "utf-16be"
          : undefined;
  const charset =
    bomCharset ?? declaredCharset?.[1] ?? declaredCharset?.[2] ?? "utf-8";

  try {
    return new TextDecoder(charset).decode(bytes);
  } catch (error) {
    if (declaredCharset || bomCharset) {
      throw new Error(`Unsupported HTML file encoding "${charset}".`, {
        cause: error,
      });
    }
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

export function normalizePastedText(text: string): string {
  return text
    .replace(/\r\n?/g, "\n")
    .replace(/[\u00a0\u202f]/g, " ")
    .replace(/\u2014/g, "-")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/\u2026/g, "...")
    .replace(/[\u00ad\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff\ufffd]/g, "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "");
}

export function normalizeTextNodes(root: ParentNode): void {
  const documentRef =
    root.nodeType === 9 ? (root as Document) : (root as Node).ownerDocument;
  if (!documentRef) {
    throw new Error("Cannot normalize text without an owning document.");
  }
  const walker = documentRef.createTreeWalker(
    root,
    documentRef.defaultView?.NodeFilter.SHOW_TEXT ?? 4,
  );
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    text.nodeValue = normalizePastedText(text.nodeValue ?? "");
  }
}

export function normalizeHeadingLevels(
  root: ParentNode,
  documentRef: Document,
): void {
  const headings = [...root.querySelectorAll("h1,h2,h3,h4,h5,h6")];
  if (!headings.length) {
    return;
  }

  const firstLevel = Math.min(
    ...headings.map((heading) => Number(heading.tagName.slice(1))),
  );
  let previousLevel = 1;
  for (const heading of headings) {
    const sourceLevel = Number(heading.tagName.slice(1));
    const level = Math.min(6, sourceLevel - firstLevel + 2, previousLevel + 1);
    const replacement = documentRef.createElement(`h${level}`);
    replacement.replaceChildren(...heading.childNodes);
    heading.replaceWith(replacement);
    previousLevel = level;
  }
}

export function sectionizeArticleHtml(
  html: string,
  documentRef: Document,
  title: string,
): ContentSection[] {
  const parsed = documentRef.implementation.createHTMLDocument("");
  parsed.body.innerHTML = html;
  const blocks = [...parsed.body.children];
  const headingSections = blocks.some((block) => block.tagName === "H2");
  const totalWords = wordCount(parsed.body.textContent ?? "");
  if (!headingSections && totalWords < 3_000) {
    return [
      { id: "section-1", title: title || "Article", words: totalWords, html },
    ];
  }

  const sections: ContentSection[] = [];
  let currentTitle = title || "Article";
  let currentBlocks: Element[] = [];
  let currentWords = 0;
  let seenHeading = false;
  const finishSection = () => {
    if (!currentBlocks.length) {
      return;
    }
    sections.push({
      id: `section-${sections.length + 1}`,
      title: currentTitle || `Section ${sections.length + 1}`,
      words: currentWords,
      html: currentBlocks.map((block) => block.outerHTML).join(""),
    });
    currentBlocks = [];
    currentWords = 0;
  };

  for (const block of blocks) {
    const blockWords = wordCount(block.textContent ?? "");
    const isHeadingBoundary = block.tagName === "H2";
    if (isHeadingBoundary) {
      if (seenHeading) {
        finishSection();
      }
      currentTitle =
        block.textContent?.replace(/\s+/g, " ").trim() || currentTitle;
      seenHeading = true;
    } else if (
      !headingSections &&
      currentWords + blockWords > 2_500 &&
      currentBlocks.length
    ) {
      finishSection();
    }
    currentBlocks.push(block);
    currentWords += blockWords;
  }
  finishSection();
  return sections;
}

function wordCount(text: string): number {
  return text.trim().match(/\S+/g)?.length ?? 0;
}
