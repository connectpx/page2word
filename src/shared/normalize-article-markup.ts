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
  label.textContent = prefix[2];
  caption.replaceChildren(
    documentRef.createTextNode(prefix[1]),
    label,
    documentRef.createTextNode(text.slice(prefix[0].length)),
  );
}
