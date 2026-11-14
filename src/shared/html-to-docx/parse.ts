import { sanitizeHtml } from "../sanitize";
import { normalizeArticleMarkup } from "../normalize-article-markup";
import type {
  DocxBlock,
  DocxParagraphBlock,
  ExportProgress,
  InlineRun,
  ParagraphStyle,
  DocxTableBlock,
  DocxTableCell,
  DocxTableRow,
} from "./types";

const blockTags = new Set([
  "ADDRESS",
  "ARTICLE",
  "ASIDE",
  "BLOCKQUOTE",
  "DIV",
  "FIGURE",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
  "HR",
  "OL",
  "P",
  "PRE",
  "SECTION",
  "TABLE",
  "UL",
]);

export async function parseHtmlToDocxBlocks(
  sections: readonly string[],
  documentRef: Document,
  onProgress?: (progress: ExportProgress) => void,
  signal?: AbortSignal,
  options: { keepOriginalLinks?: boolean } = {},
): Promise<DocxBlock[]> {
  const blocks: DocxBlock[] = [];
  let activeSection = 0;
  let lastYield = performance.now();
  let listInstance = 0;

  for (
    let sectionIndex = 0;
    sectionIndex < sections.length;
    sectionIndex += 1
  ) {
    activeSection = sectionIndex;
    throwIfAborted(signal);
    const parsed = documentRef.implementation.createHTMLDocument("");
    parsed.body.innerHTML = sanitizeHtml(sections[sectionIndex], documentRef);
    if (!options.keepOriginalLinks) {
      removeOriginalLinks(parsed.body);
    }
    normalizeArticleMarkup(parsed.body, documentRef);
    const fragment = parsed.createDocumentFragment();
    while (parsed.body.firstChild) {
      fragment.append(parsed.body.firstChild);
    }

    for (const child of [...fragment.childNodes]) {
      await walkBlock(child, "body", 0);
    }
    onProgress?.({
      stage: "preparing",
      completed: sectionIndex + 1,
      total: sections.length,
    });
  }

  return blocks;

  async function walkBlock(
    node: Node,
    style: ParagraphStyle,
    listLevel: number,
    listType?: "bullet" | "number",
    listId?: number,
  ): Promise<void> {
    throwIfAborted(signal);
    if (node.nodeType === 3) {
      const runs = inlineRuns(node);
      if (hasText(runs)) {
        addParagraph(runs, style, listType, listLevel, listId);
      }
      return;
    }
    if (node.nodeType !== 1) {
      return;
    }

    const element = node as Element;
    const tag = element.tagName;
    if (tag === "UL" || tag === "OL") {
      const nextId = ++listInstance;
      const type = tag === "UL" ? "bullet" : "number";
      for (const item of [...element.children]) {
        if (item.tagName !== "LI") {
          await walkBlock(item, style, listLevel, type, nextId);
          continue;
        }
        const inlineNodes = [...item.childNodes].filter(
          (child) =>
            !(
              isElementNode(child) &&
              (child.tagName === "UL" || child.tagName === "OL")
            ),
        );
        const runs = inlineNodes.flatMap((child) => inlineRuns(child));
        if (hasText(runs)) {
          addParagraph(runs, style, type, Math.min(listLevel, 8), nextId);
        }
        for (const nested of [...item.children].filter(
          (child) => child.tagName === "UL" || child.tagName === "OL",
        )) {
          await walkBlock(nested, style, listLevel + 1);
        }
      }
      await yieldIfNeeded();
      return;
    }
    if (/^H[1-6]$/.test(tag)) {
      const level = Math.min(Number(tag.slice(1)), 4);
      addParagraph(inlineRuns(element), `heading${level}` as ParagraphStyle);
      await yieldIfNeeded();
      return;
    }
    if (tag === "P") {
      const runs = inlineRuns(element);
      if (hasText(runs)) {
        const paragraphStyle =
          element.querySelector("img") && !element.textContent?.trim()
            ? "image"
            : style;
        addParagraph(
          runs,
          paragraphStyle,
          listType,
          listLevel,
          listId,
          paragraphStyle === "image"
            ? validBookmark(element.querySelector("img")?.id)
            : undefined,
        );
      }
      await yieldIfNeeded();
      return;
    }
    if (tag === "PRE") {
      const lines = (element.textContent ?? "")
        .replace(/\r\n?/g, "\n")
        .split("\n");
      for (const line of lines) {
        addParagraph([{ text: line || " " }], "code");
      }
      await yieldIfNeeded();
      return;
    }
    if (tag === "HR") {
      addParagraph([], "rule");
      await yieldIfNeeded();
      return;
    }
    if (tag === "IMG" || tag === "FIGURE") {
      const image = tag === "IMG" ? element : element.querySelector("img");
      const captionElement =
        tag === "FIGURE" ? figureCaptionElement(element) : null;
      const caption = captionElement?.textContent?.trim();
      if (image) {
        addParagraph(
          [imageInlineRun(image, caption)],
          "image",
          undefined,
          0,
          undefined,
          caption ? undefined : validBookmark(element.id),
        );
      }
      if (captionElement && caption) {
        addParagraph(
          trimRuns(inlineRuns(captionElement)),
          "figureCaption",
          undefined,
          0,
          undefined,
          validBookmark(element.id),
        );
      }
      await yieldIfNeeded();
      return;
    }
    if (tag === "TABLE") {
      const caption = element
        .querySelector(":scope > caption")
        ?.textContent?.trim();
      const table = await parseTable(element, yieldIfNeeded);
      if (caption) {
        addParagraph([{ text: caption }], "caption");
      }
      if (table.rows.length && table.columnCount) {
        table.bookmark = validBookmark(element.id);
        blocks.push(table);
      }
      return;
    }
    if (tag === "BLOCKQUOTE") {
      for (const child of [...element.childNodes]) {
        if (isElementNode(child) && blockTags.has(child.tagName)) {
          await walkBlock(child, "quote", listLevel);
        } else {
          const runs = inlineRuns(child);
          if (hasText(runs)) {
            addParagraph(runs, "quote", listType, listLevel, listId);
          }
        }
      }
      await yieldIfNeeded();
      return;
    }

    const inlineStyle =
      element.querySelector("img") && !element.textContent?.trim()
        ? "image"
        : style;
    if (!element.children.length) {
      const runs = inlineRuns(element);
      if (hasText(runs)) {
        addParagraph(runs, inlineStyle, listType, listLevel, listId);
      }
      return;
    }
    for (const child of [...element.childNodes]) {
      if (isElementNode(child) && blockTags.has(child.tagName)) {
        await walkBlock(child, inlineStyle, listLevel, listType, listId);
      } else {
        const runs = inlineRuns(child);
        if (hasText(runs)) {
          addParagraph(runs, inlineStyle, listType, listLevel, listId);
        }
      }
    }
    await yieldIfNeeded();
  }

  function addParagraph(
    runs: InlineRun[],
    style: ParagraphStyle,
    listType?: "bullet" | "number",
    listLevel = 0,
    listId?: number,
    bookmark?: string,
  ): void {
    const cleanRuns = trimRuns(runs);
    if (!cleanRuns.length && style !== "code" && style !== "rule") {
      return;
    }
    const block: DocxParagraphBlock = {
      kind: "paragraph",
      style,
      runs: cleanRuns,
      ...(bookmark ? { bookmark } : {}),
    };
    if (listType && listId !== undefined) {
      block.list = { type: listType, level: listLevel, instance: listId };
    }
    blocks.push(block);
  }

  async function yieldIfNeeded(): Promise<void> {
    onProgress?.({
      stage: "preparing",
      completed: activeSection,
      total: sections.length,
    });
    if (performance.now() - lastYield < 25) {
      return;
    }
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    lastYield = performance.now();
    throwIfAborted(signal);
  }
}

function inlineRuns(
  node: Node,
  marks: Omit<InlineRun, "text"> = {},
): InlineRun[] {
  if (node.nodeType === 3) {
    const text = (node.nodeValue ?? "").replace(/[\t\n\r\f ]+/g, " ");
    return text ? [{ ...marks, text }] : [];
  }
  if (node.nodeType !== 1) {
    return [];
  }

  const element = node as Element;
  const tag = element.tagName;
  if (tag === "BR") {
    return [{ ...marks, text: "", lineBreak: true }];
  }
  if (tag === "IMG") {
    const image = imageInlineRun(element);
    return [{ ...marks, ...image, href: marks.href, anchor: marks.anchor }];
  }
  const next = { ...marks };
  if (tag === "STRONG" || tag === "B") next.bold = true;
  if (tag === "EM" || tag === "I") next.italic = true;
  if (tag === "U") next.underline = true;
  if (tag === "S" || tag === "DEL") next.strike = true;
  if (tag === "SUB") next.subscript = true;
  if (tag === "SUP") next.superscript = true;
  if (tag === "CODE") next.code = true;
  if (tag === "A") {
    const href = element.getAttribute("href");
    if (href && /^(?:https?:|mailto:|tel:)/i.test(href)) {
      next.href = href;
    } else if (href && /^#[\w-]+$/.test(href)) {
      next.anchor = href.slice(1);
    }
  }

  return [...element.childNodes].flatMap((child) => inlineRuns(child, next));
}

function removeOriginalLinks(root: ParentNode): void {
  for (const link of root.querySelectorAll<HTMLAnchorElement>("a[href]")) {
    if (
      /^#p2w_(?:figure|table|image)_/i.test(link.getAttribute("href") ?? "")
    ) {
      continue;
    }
    link.replaceWith(...link.childNodes);
  }
}

function validBookmark(value: string | null | undefined): string | undefined {
  return value && /^p2w_(?:figure|table|image)_\d+(?:_\d+)?$/.test(value)
    ? value
    : undefined;
}

function trimRuns(runs: InlineRun[]): InlineRun[] {
  const result = runs.filter((run) => run.text || run.lineBreak || run.assetId);
  const firstText = result.find((run) => run.text);
  const lastText = [...result].reverse().find((run) => run.text);
  if (firstText) firstText.text = firstText.text.replace(/^\s+/, "");
  if (lastText) lastText.text = lastText.text.replace(/\s+$/, "");
  return result.filter((run) => run.text || run.lineBreak || run.assetId);
}

function hasText(runs: InlineRun[]): boolean {
  return runs.some((run) => run.text.trim() || run.lineBreak || run.assetId);
}

function figureCaptionElement(figure: Element): Element | null {
  return (
    [...figure.querySelectorAll("figcaption")].find(
      (candidate) => candidate.closest("figure") === figure,
    ) ??
    [...figure.querySelectorAll("p")].find(
      (candidate) =>
        candidate.closest("figure") === figure &&
        /^fig(?:ure)?\.?\s*\d+\b/i.test(candidate.textContent?.trim() ?? ""),
    ) ??
    null
  );
}

async function parseTable(
  table: Element,
  yieldToEventLoop: () => Promise<void>,
): Promise<DocxTableBlock> {
  const rows = [...table.querySelectorAll("tr")].filter(
    (row) => row.closest("table") === table,
  );
  const occupied: Array<Array<DocxTableCell | undefined>> = [];
  const tableRows: DocxTableRow[] = [];
  let columnCount = 0;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const rowHeader = row.parentElement?.tagName === "THEAD";
    const cells: DocxTableCell[] = [];
    let columnIndex = 0;
    for (const cellElement of [
      ...row.querySelectorAll(":scope > th, :scope > td"),
    ]) {
      while (occupied[rowIndex]?.[columnIndex]) {
        columnIndex += 1;
      }
      const columnSpan = Math.min(
        50,
        Math.max(1, positiveInteger(cellElement.getAttribute("colspan"))),
      );
      const rowSpan = Math.min(
        rows.length - rowIndex,
        Math.max(1, positiveInteger(cellElement.getAttribute("rowspan"))),
      );
      const header = rowHeader || cellElement.tagName === "TH";
      const cell: DocxTableCell = {
        runs: trimRuns(inlineRuns(cellElement)),
        columnSpan,
        rowSpan,
        header,
      };
      cells.push(cell);
      for (
        let mergeRow = rowIndex;
        mergeRow < rowIndex + rowSpan;
        mergeRow += 1
      ) {
        occupied[mergeRow] ??= [];
        for (
          let mergeColumn = columnIndex;
          mergeColumn < columnIndex + columnSpan;
          mergeColumn += 1
        ) {
          occupied[mergeRow][mergeColumn] = cell;
        }
      }
      columnIndex += columnSpan;
    }
    columnCount = Math.max(columnCount, columnIndex);
    tableRows.push({
      cells,
      header: rowHeader || cells.some((cell) => cell.header),
    });
    await yieldToEventLoop();
  }

  return { kind: "table", rows: tableRows, columnCount };
}

function imageInlineRun(element: Element, fallbackAlt?: string): InlineRun {
  const assetId = element.getAttribute("data-p2w-asset") ?? "";
  const alt = element.getAttribute("alt")?.trim() || fallbackAlt || "image";
  if (!assetId) {
    return {
      text: `[Image not embedded in this export: ${alt}]`,
      italic: true,
    };
  }
  return {
    text: "",
    assetId,
    alt,
    width: positiveInteger(element.getAttribute("width")),
    height: positiveInteger(element.getAttribute("height")),
  };
}

function positiveInteger(value: string | null): number {
  const parsed = Number.parseInt(value ?? "", 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : 0;
}

function isElementNode(node: Node): node is Element {
  return node.nodeType === 1;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) {
    throw new DOMException("Export cancelled.", "AbortError");
  }
}
