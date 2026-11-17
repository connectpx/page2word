import { Readability } from "@mozilla/readability";
import type { ExtractedDoc } from "../../shared/types";
import {
  normalizeHeadingLevels,
  normalizeTextNodes,
  sectionizeArticleHtml,
} from "../../shared/html-document";
import { sanitizeHtml } from "../../shared/sanitize";

interface Article {
  title: string;
  byline: string | null;
  siteName: string | null;
  lang: string | null;
  content: string;
  textContent: string;
}

function wordCount(text: string): number {
  return text.trim().match(/\S+/g)?.length ?? 0;
}

function createId(): string {
  return (
    globalThis.crypto?.randomUUID?.() ??
    `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
  );
}

function bestSrcsetCandidate(srcset: string): string | null {
  const candidates = srcset
    .split(",")
    .map((candidate) => {
      const [url, descriptor] = candidate.trim().split(/\s+/, 2);
      const amount = Number.parseFloat(descriptor ?? "1");
      const score = descriptor?.endsWith("w") ? amount : amount * 1000;
      return { url, score };
    })
    .filter((candidate) => candidate.url);
  return (
    candidates.sort((left, right) => right.score - left.score)[0]?.url ?? null
  );
}

function absoluteUrl(value: string | null, baseUrl: string): string | null {
  if (!value) {
    return null;
  }
  try {
    const url = new URL(value, baseUrl);
    return ["http:", "https:", "data:"].includes(url.protocol)
      ? url.href
      : null;
  } catch {
    return null;
  }
}

function prepareDocument(source: Document): void {
  for (const image of source.querySelectorAll("img")) {
    const currentSrc = image.getAttribute("src")?.trim() ?? "";
    const needsLazySource =
      !currentSrc || /^(?:about:blank|data:image\/gif)/i.test(currentSrc);
    const pictureSources = [
      ...(image.closest("picture")?.querySelectorAll("source") ?? []),
    ];
    const lazyCandidates = [
      image.getAttribute("data-src"),
      image.getAttribute("data-original"),
      image.getAttribute("data-lazy-src"),
      bestSrcsetCandidate(image.getAttribute("data-srcset") ?? ""),
      ...pictureSources.map((source) =>
        bestSrcsetCandidate(source.getAttribute("srcset") ?? ""),
      ),
      ...pictureSources.map((source) => source.getAttribute("src")),
      bestSrcsetCandidate(image.getAttribute("srcset") ?? ""),
    ];
    const selected = needsLazySource
      ? (lazyCandidates.find(Boolean) ?? null)
      : currentSrc;
    const src = absoluteUrl(selected, source.baseURI);
    if (src) {
      image.setAttribute("src", src);
    } else {
      image.removeAttribute("src");
    }
    const width =
      image.naturalWidth || positiveDimension(image.getAttribute("width"));
    const height =
      image.naturalHeight || positiveDimension(image.getAttribute("height"));
    if (width) {
      image.setAttribute("width", String(width));
    }
    if (height) {
      image.setAttribute("height", String(height));
    }
  }

  for (const link of source.querySelectorAll("a[href]")) {
    const href = absoluteUrl(link.getAttribute("href"), source.baseURI);
    if (href && /^(?:https?:|mailto:|tel:)/i.test(href)) {
      link.setAttribute("href", href);
    } else {
      link.removeAttribute("href");
    }
  }
}

function positiveDimension(value: string | null): number {
  const dimension = Number.parseInt(value ?? "", 10);
  return Number.isFinite(dimension) && dimension > 0 ? dimension : 0;
}

function normalizeContent(
  html: string,
  documentRef: Document,
  title: string,
): string {
  const parsed = documentRef.implementation.createHTMLDocument("");
  parsed.body.innerHTML = html;
  const leadingHeading = parsed.body.firstElementChild;
  if (
    leadingHeading &&
    /^H[1-6]$/.test(leadingHeading.tagName) &&
    leadingHeading.textContent?.trim().toLowerCase() === title.toLowerCase()
  ) {
    leadingHeading.remove();
  }

  for (const container of [
    ...parsed.body.querySelectorAll("div,span"),
  ].reverse()) {
    if (!container.parentNode) {
      continue;
    }
    if (
      container.tagName === "DIV" &&
      !container.querySelector(
        "address,article,aside,blockquote,div,dl,fieldset,figure,footer,form,h1,h2,h3,h4,h5,h6,header,hr,main,ol,p,pre,section,table,ul",
      )
    ) {
      const paragraph = parsed.createElement("p");
      paragraph.replaceChildren(...container.childNodes);
      container.replaceWith(paragraph);
    } else {
      container.replaceWith(...container.childNodes);
    }
  }

  normalizeTextNodes(parsed.body);
  normalizeHeadingLevels(parsed.body, parsed);

  return sanitizeHtml(parsed.body.innerHTML, documentRef);
}

function findAbstract(source: Document): Element | null {
  const abstractSelector =
    '[id*="abstract" i],[class*="abstract" i],[data-section*="abstract" i]';
  const heading = [...source.querySelectorAll("h1,h2,h3,h4,h5,h6")].find(
    (element) => /^abstract\b/i.test(element.textContent?.trim() ?? ""),
  );
  if (heading) {
    return heading.closest(abstractSelector) ?? heading.parentElement;
  }

  return (
    [...source.querySelectorAll(abstractSelector)]
      .filter((element) => wordCount(element.textContent ?? "") >= 10)
      .sort(
        (left, right) =>
          (right.textContent?.length ?? 0) - (left.textContent?.length ?? 0),
      )[0] ?? null
  );
}

function recoverMissingAbstract(
  source: Document,
  articleText: string,
  title: string,
): string | null {
  const candidate = findAbstract(source);
  if (!candidate) {
    return null;
  }

  const abstractText =
    candidate.textContent?.replace(/^abstract\b/i, "").trim() ?? "";
  if (wordCount(abstractText) < 10) {
    return null;
  }
  const normalizedArticle = articleText.replace(/\s+/g, " ").toLowerCase();
  const normalizedAbstract = abstractText.replace(/\s+/g, " ").toLowerCase();
  if (
    normalizedAbstract.length >= 40 &&
    normalizedArticle.includes(normalizedAbstract)
  ) {
    return null;
  }

  const html = normalizeContent(candidate.innerHTML, source, title);
  return html || null;
}

function fallbackArticle(source: Document): Article | null {
  const candidates = source.querySelectorAll("article,main,[role=main]");
  let best: Element | null = null;
  let bestScore = 0;

  for (const candidate of candidates) {
    const text = candidate.textContent?.trim() ?? "";
    const score = wordCount(text) + candidate.querySelectorAll("p").length * 20;
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  }

  if (!best) {
    return null;
  }

  const clone = best.cloneNode(true) as Element;
  clone
    .querySelectorAll(
      "nav,aside,footer,header,form,[role=banner],[role=navigation],[aria-label*=comment i],.ad,.advert,.related,.share,.newsletter",
    )
    .forEach((element) => element.remove());
  const content = clone.innerHTML;
  return {
    title: source.title || "Untitled article",
    byline: null,
    siteName: source.location.hostname,
    lang: source.documentElement.lang || null,
    content,
    textContent: clone.textContent ?? "",
  };
}

export function extractArticle(source: Document): ExtractedDoc {
  const warnings: string[] = [];
  const clone = source.cloneNode(true) as Document;
  prepareDocument(clone);
  const parsed = new Readability(clone, {
    keepClasses: false,
    charThreshold: 200,
  }).parse() as Article | null;
  let article = parsed;
  const readabilityWords = article ? wordCount(article.textContent) : 0;
  if (!article || readabilityWords < 500) {
    const fallback = fallbackArticle(source);
    const fallbackWords = fallback ? wordCount(fallback.textContent) : 0;
    const fallbackIsSubstantiallyLonger =
      fallbackWords - readabilityWords >= Math.max(50, readabilityWords * 0.5);
    if (
      fallback &&
      (!article || readabilityWords < 30 || fallbackIsSubstantiallyLonger)
    ) {
      article = fallback;
      warnings.push(
        "Reader mode omitted substantial article text; a best-effort page-area extraction was used.",
      );
    }
  }
  if (!article) {
    throw new Error("Could not find article content on this page.");
  }

  const title = article.title.trim() || source.title || "Untitled article";
  const normalizedContent = normalizeContent(article.content, source, title);
  const recoveredAbstract = recoverMissingAbstract(
    source,
    article.textContent,
    title,
  );
  const html = recoveredAbstract
    ? `${recoveredAbstract}${normalizedContent}`
    : normalizedContent;
  if (recoveredAbstract) {
    warnings.push(
      "A labeled abstract section omitted by reader mode was restored.",
    );
  }
  const parsedContent = source.implementation.createHTMLDocument("");
  parsedContent.body.innerHTML = html;
  const text = parsedContent.body.textContent ?? "";
  const contentSections = sectionizeArticleHtml(html, source, title);

  return {
    id: createId(),
    title,
    ...(article.byline ? { byline: article.byline } : {}),
    ...(article.siteName ? { siteName: article.siteName } : {}),
    sourceUrl: source.location.href,
    extractedAt: new Date().toISOString(),
    ...(article.lang ? { lang: article.lang } : {}),
    html,
    sections: contentSections.map(({ id, title: sectionTitle, words }) => ({
      id,
      title: sectionTitle,
      words,
    })),
    contentSections,
    assets: [],
    stats: {
      words: wordCount(text),
      images: parsedContent.body.querySelectorAll("img").length,
      tables: parsedContent.body.querySelectorAll("table").length,
      charts: parsedContent.body.querySelectorAll("figure").length,
    },
    warnings,
  };
}

declare global {
  interface Window {
    page2wordExtract?: () => ExtractedDoc;
  }
}

if (typeof window !== "undefined" && typeof document !== "undefined") {
  window.page2wordExtract = () => extractArticle(document);
}
