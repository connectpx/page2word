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
