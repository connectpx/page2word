export function normalizeArticleMarkup(
  root: ParentNode,
  documentRef: Document,
): void {
  for (const item of root.querySelectorAll("ul li")) {
    if (item.closest("ul,ol")?.tagName !== "UL") {
      continue;
    }
    const text = firstListText(item);
    if (text) {
      text.nodeValue = (text.nodeValue ?? "").replace(/^\s*•(?:\s+|$)/, "");
    }
  }

  for (const figure of root.querySelectorAll("figure")) {
    const caption =
      [...figure.querySelectorAll("figcaption")].find(
        (candidate) => candidate.closest("figure") === figure,
      ) ??
      [...figure.querySelectorAll("p")].find(
        (candidate) =>
          candidate.closest("figure") === figure &&
          /^fig(?:ure)?\.?\s*\d+\b/i.test(candidate.textContent?.trim() ?? ""),
      );
    if (!caption) {
      continue;
    }

    let figureCaption = caption;
    if (caption.tagName !== "FIGCAPTION") {
      figureCaption = documentRef.createElement("figcaption");
      figureCaption.replaceChildren(...caption.childNodes);
      caption.replaceWith(figureCaption);
    }
    formatFigureLabel(figureCaption, documentRef);
  }

  for (const table of root.querySelectorAll("table")) {
    const caption =
      table.querySelector(":scope > caption") ??
      findAdjacentTableCaption(table);
    if (caption) {
      formatTableLabel(caption, documentRef);
    }
    normalizeTableHeader(table, documentRef);
  }

  normalizeFigureReferences(root, documentRef);
}

function firstListText(item: Element): Text | undefined {
  for (const child of item.childNodes) {
    if (child.nodeType === 3) {
      if (child.textContent?.trim()) {
        return child as Text;
      }
      continue;
    }
    if (child.nodeType !== 1) {
      continue;
    }
    const element = child as Element;
    if (element.tagName === "UL" || element.tagName === "OL") {
      continue;
    }
    const text = firstListText(element);
    if (text) {
      return text;
    }
  }
  return undefined;
}

function formatFigureLabel(caption: Element, documentRef: Document): void {
  const text = caption.textContent ?? "";
  const prefix = text.match(/^(\s*)(fig(?:ure)?\.?)(?=\s*\d)/i);
  if (!prefix) {
    return;
  }

  const label = documentRef.createElement("strong");
  label.textContent = "Figure";
  caption.replaceChildren(
    documentRef.createTextNode(prefix[1]),
    label,
    documentRef.createTextNode(text.slice(prefix[0].length)),
  );
}

function normalizeFigureReferences(
  root: ParentNode,
  documentRef: Document,
): void {
  const walker = documentRef.createTreeWalker(
    root,
    documentRef.defaultView?.NodeFilter.SHOW_TEXT ?? 4,
  );
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    text.nodeValue =
      text.nodeValue?.replace(/\bFig\.?(?=\s*\d)/gi, "Figure") ?? "";
  }
}

function formatTableLabel(caption: Element, documentRef: Document): void {
  caption.classList.add("p2w-table-caption");
  const firstText = firstCaptionText(caption);
  const match = firstText?.nodeValue?.match(/^(\s*Table\s+\d+\.?)(?=\s|$)/i);
  let label = firstText?.parentElement?.closest("strong,b") ?? null;

  if (match && firstText && !label) {
    label = documentRef.createElement("strong");
    label.textContent = match[1];
    const remainder = firstText.nodeValue?.slice(match[1].length) ?? "";
    const fragment = documentRef.createDocumentFragment();
    fragment.append(label);
    if (remainder) {
      fragment.append(documentRef.createTextNode(remainder));
    }
    firstText.replaceWith(fragment);
  }

  if (label) {
    const link = label.closest("a");
    if (link) {
      link.replaceWith(...link.childNodes);
    }
  }
}

function findAdjacentTableCaption(table: Element): Element | null {
  let previous = table.previousElementSibling;
  while (previous && previous.tagName !== "P") {
    if (previous.tagName === "TABLE") {
      return null;
    }
    previous = previous.previousElementSibling;
  }
  return previous && /^Table\s+\d+\b/i.test(previous.textContent?.trim() ?? "")
    ? previous
    : null;
}

function normalizeTableHeader(table: Element, documentRef: Document): void {
  if (table.querySelector("thead,th")) {
    return;
  }
  const firstRow = table.querySelector("tr");
  if (!firstRow) {
    return;
  }
  for (const cell of [...firstRow.querySelectorAll(":scope > td")]) {
    const header = documentRef.createElement("th");
    for (const attribute of [...cell.attributes]) {
      header.setAttribute(attribute.name, attribute.value);
    }
    header.replaceChildren(...cell.childNodes);
    cell.replaceWith(header);
  }
}

function firstCaptionText(root: Element): Text | undefined {
  const walker = root.ownerDocument.createTreeWalker(
    root,
    root.ownerDocument.defaultView?.NodeFilter.SHOW_TEXT ?? 4,
  );
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (text.nodeValue?.trim()) {
      return text;
    }
  }
  return undefined;
}
