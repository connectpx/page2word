import createDOMPurify from "dompurify";

const allowedTags = [
  "a",
  "b",
  "blockquote",
  "br",
  "caption",
  "code",
  "col",
  "colgroup",
  "em",
  "figcaption",
  "figure",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "hr",
  "i",
  "img",
  "li",
  "ol",
  "p",
  "pre",
  "s",
  "span",
  "strong",
  "sub",
  "sup",
  "table",
  "tbody",
  "td",
  "tfoot",
  "th",
  "thead",
  "tr",
  "u",
  "ul",
];
const allowedAttributes = [
  "alt",
  "colspan",
  "data-p2w-asset",
  "height",
  "href",
  "id",
  "rowspan",
  "src",
  "style",
  "title",
  "width",
  "data-p2w-original-href",
];
const junkElementSelector = [
  "grammarly-extension",
  "grammarly-desktop-integration",
  "[data-grammarly-shadow-root]",
  "quillbot-extension",
  "lt-toolbar",
  "[data-new-gr-c-s-check-loaded]",
].join(",");

export function sanitizeHtml(
  html: string,
  documentRef: Document = document,
): string {
  const view = documentRef.defaultView;
  if (!view) {
    throw new Error("Cannot sanitize HTML without a live document window.");
  }

  const parsed = documentRef.implementation.createHTMLDocument("");
  parsed.body.innerHTML = html;

  for (const element of parsed.body.querySelectorAll(junkElementSelector)) {
    element.remove();
  }

  for (const element of parsed.body.querySelectorAll("*")) {
    const rawStyle = element.getAttribute("style");
    if (rawStyle) {
      const style = parsed.createElement("span").style;
      style.cssText = rawStyle;
      const safeStyle = parsed.createElement("span").style;
      const fontFamily = style.getPropertyValue("font-family").trim();
      const fontSize = style.getPropertyValue("font-size").trim();
      const fontWeight = style.getPropertyValue("font-weight").trim();
      const color = style.getPropertyValue("color").trim();
      const textAlign = style.getPropertyValue("text-align").trim();
      const padding = style.getPropertyValue("padding").trim();
      const border = style.getPropertyValue("border").trim();
      if (/^[\w\s,'"-]{1,80}$/.test(fontFamily)) {
        safeStyle.setProperty("font-family", fontFamily);
      }
      if (/^(?:\d+(?:\.\d+)?)(?:px|pt)$/i.test(fontSize)) {
        safeStyle.setProperty("font-size", fontSize);
      }
      if (/^(?:normal|bold|[1-9]00)$/i.test(fontWeight)) {
        safeStyle.setProperty("font-weight", fontWeight);
      }
      if (isSafeTextColor(color)) {
        safeStyle.setProperty("color", color);
      }
      if (
        textAlign === "left" ||
        textAlign === "center" ||
        textAlign === "right"
      ) {
        safeStyle.setProperty("text-align", textAlign);
      }
      if (/^(?:\d+(?:\.\d+)?)(?:px)$/i.test(padding)) {
        safeStyle.setProperty("padding", padding);
      }
      if (/^1px solid (?:#202124|rgb\(32,\s*33,\s*36\))$/i.test(border)) {
        safeStyle.setProperty("border", border);
      }
      if (safeStyle.cssText) {
        element.setAttribute("style", safeStyle.cssText);
      } else {
        element.removeAttribute("style");
      }
    }
    const hasExtensionMarkup =
      element.tagName.toLowerCase() === "lt-highlighter" ||
      [...element.attributes].some(
        ({ name, value }) =>
          name.startsWith("data-lt-") ||
          name.startsWith("data-gr-") ||
          name.startsWith("data-gramm") ||
          name.startsWith("data-qb-") ||
          (name === "class" && /\b(?:gr_|qb-)/i.test(value)) ||
          (name === "id" && value.startsWith("qb-")),
      );
    if (hasExtensionMarkup && element.parentNode) {
      element.replaceWith(...element.childNodes);
    }
  }

  const purifier = createDOMPurify(view);
  const clean = purifier.sanitize(parsed.body.innerHTML, {
    ALLOWED_TAGS: allowedTags,
    ALLOWED_ATTR: allowedAttributes,
    ADD_ATTR: ["colspan", "rowspan", "data-p2w-asset"],
    ADD_URI_SAFE_ATTR: [
      "colspan",
      "rowspan",
      "data-p2w-asset",
      "data-p2w-original-href",
      "height",
      "width",
    ],
    ALLOW_DATA_ATTR: false,
    ALLOWED_URI_REGEXP:
      /^(?:(?:https?|mailto|tel):|#[\w-]+|data:image\/(?:png|jpeg|gif);base64,|p2w-asset:|blob:)/i,
  });
  const result = documentRef.implementation.createHTMLDocument("");
  result.body.innerHTML = clean;

  for (const element of result.body.querySelectorAll(
    "p,li,blockquote,figcaption",
  )) {
    if (!element.textContent?.trim() && !element.querySelector("img,br")) {
      element.remove();
    }
  }
  for (const image of result.body.querySelectorAll("img")) {
    const src = image.getAttribute("src") ?? "";
    if (
      !/^(?:https?:|data:image\/(?:png|jpeg|gif);base64,|p2w-asset:|blob:)/i.test(
        src,
      )
    ) {
      image.remove();
    }
  }
  for (const link of result.body.querySelectorAll("a")) {
    const href = link.getAttribute("href") ?? "";
    if (!/^(?:https?:|mailto:|tel:|#[\w-]+$)/i.test(href)) {
      link.removeAttribute("href");
    }
  }

  return result.body.innerHTML;
}

function isSafeTextColor(value: string): boolean {
  if (/^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(value)) {
    return true;
  }
  const channels = value.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/i);
  return Boolean(
    channels &&
    channels
      .slice(1)
      .every((channel) => Number(channel) >= 0 && Number(channel) <= 255),
  );
}
