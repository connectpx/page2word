import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import { normalizeArticleMarkup } from "../src/shared/normalize-article-markup";

describe("article markup normalization", () => {
  it("removes a literal bullet only when a UL already supplies the marker", () => {
    const dom = new JSDOM("");
    const article = dom.window.document.createElement("article");
    article.innerHTML =
      "<ul><li><strong>• </strong>First<ul><li>• Second</li></ul></li></ul>" +
      "<ol><li>• Keep ordered text</li></ol>";

    normalizeArticleMarkup(article, dom.window.document);

    const unorderedItems = [...article.querySelectorAll("ul li")];
    expect(unorderedItems.map((item) => item.textContent?.trim())).toEqual([
      "FirstSecond",
      "Second",
    ]);
    expect(article.querySelector("ol li")?.textContent?.trim()).toBe(
      "• Keep ordered text",
    );
    dom.window.close();
  });

  it("preserves and bolds the original figure label", () => {
    const dom = new JSDOM("");
    const article = dom.window.document.createElement("article");
    article.innerHTML =
      "<figure><p><strong>Fig.</strong> 2. Caption text</p></figure>";

    normalizeArticleMarkup(article, dom.window.document);
    normalizeArticleMarkup(article, dom.window.document);

    const caption = article.querySelector("figure figcaption");
    expect(caption?.textContent).toBe("Fig. 2. Caption text");
    expect(caption?.querySelector("strong")?.textContent).toBe("Fig.");
    expect(caption?.innerHTML).toBe("<strong>Fig.</strong> 2. Caption text");
    dom.window.close();
  });
});
