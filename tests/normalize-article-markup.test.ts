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

  it("normalizes and bolds figure captions and references", () => {
    const dom = new JSDOM("");
    const article = dom.window.document.createElement("article");
    article.innerHTML =
      "<figure><p><strong>Fig.</strong> 2. Caption text</p></figure>" +
      "<p>See Fig. 2 and Figure 3.</p>";

    normalizeArticleMarkup(article, dom.window.document);
    normalizeArticleMarkup(article, dom.window.document);

    const caption = article.querySelector("figure figcaption");
    expect(caption?.textContent).toBe("Figure 2. Caption text");
    expect(caption?.querySelector("strong")?.textContent).toBe("Figure");
    expect(caption?.innerHTML).toBe("<strong>Figure</strong> 2. Caption text");
    expect(article.querySelector("p:not(figure p)")?.textContent).toBe(
      "See Figure 2 and Figure 3.",
    );
    dom.window.close();
  });

  it("bolds but does not link a table caption label", () => {
    const dom = new JSDOM("");
    const article = dom.window.document.createElement("article");
    article.innerHTML =
      '<table id="p2w_table_1"><caption>Table 1. Summary of findings.</caption>' +
      "<tr><td>Finding</td></tr></table>";

    normalizeArticleMarkup(article, dom.window.document);
    normalizeArticleMarkup(article, dom.window.document);

    expect(article.querySelector("caption")?.innerHTML).toBe(
      "<strong>Table 1.</strong> Summary of findings.",
    );
    expect(
      article.querySelector("caption")?.classList.contains("p2w-table-caption"),
    ).toBe(true);
    dom.window.close();
  });

  it("handles Word-style captions outside tables and promotes first-row cells", () => {
    const dom = new JSDOM("");
    const article = dom.window.document.createElement("article");
    article.innerHTML =
      "<p><b>Table 1.</b> Summary of key findings.</p>" +
      '<table id="p2w_table_1"><tbody><tr><td>Finding</td></tr>' +
      "<tr><td>Result</td></tr></tbody></table>";

    normalizeArticleMarkup(article, dom.window.document);
    expect(article.querySelector("p")?.innerHTML).toBe(
      "<b>Table 1.</b> Summary of key findings.",
    );
    expect(article.querySelector("table tr:first-child th")?.textContent).toBe(
      "Finding",
    );
    expect(article.querySelector("table tr:last-child td")?.textContent).toBe(
      "Result",
    );
    dom.window.close();
  });
});
