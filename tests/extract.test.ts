import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";
import { describe, expect, it } from "vitest";
import { extractArticle } from "../src/extension/extractor/extract";
import { createLargeFixture } from "./fixtures/generate-large.mjs";

const fixturesDir = resolve(
  dirname(fileURLToPath(import.meta.url)),
  "fixtures",
);

describe("article extraction", () => {
  it.each([
    "news",
    "blog",
    "documentation",
    "wikipedia",
    "table",
    "publisher-abstract",
  ])("extracts the main content from the %s fixture", (fixture) => {
    const html = readFileSync(resolve(fixturesDir, `${fixture}.html`), "utf8");
    const dom = new JSDOM(html, { url: `https://example.test/${fixture}` });
    const result = extractArticle(dom.window.document);

    expect(result.title).not.toBe("");
    expect(result.stats.words).toBeGreaterThan(30);
    expect(result.html).not.toMatch(
      /World Politics|Site menu|Documentation navigation|Donate|External reference list/,
    );
    if (fixture !== "publisher-abstract") {
      expect(result.warnings).toEqual([]);
    }
    if (fixture === "documentation") {
      expect(result.html).toContain("<pre><code>");
    }
    if (fixture === "blog") {
      expect(result.html).toContain(
        'href="https://example.test/guide/details"',
      );
      expect(result.html).toContain(
        'src="https://example.test/assets/desk-photo.jpg"',
      );
      expect(result.html).toContain("<h3>Finishing</h3>");
      expect(result.html).toContain("<strong>careful</strong>");
      expect(result.html).toContain("<em>accurate</em>");
      expect(result.html).toContain("<li>Measure each board</li>");
    }
    if (fixture === "wikipedia" || fixture === "table") {
      expect(result.stats.tables).toBeGreaterThan(0);
    }
    if (fixture === "publisher-abstract") {
      expect(result.html).toContain("ABSTRACT-ONLY-TEXT");
      expect(result.html).toContain("Introduction");
      const headings = [
        ...new JSDOM(result.html).window.document.querySelectorAll(
          "h1,h2,h3,h4,h5,h6",
        ),
      ].map((heading) => heading.textContent?.trim());
      expect(headings.indexOf("Abstract")).toBeGreaterThanOrEqual(0);
      expect(result.warnings).toContain(
        "Reader mode omitted substantial article text; a best-effort page-area extraction was used.",
      );
    }
    dom.window.close();
  });

  it("uses a best-effort article fallback for short content", () => {
    const dom = new JSDOM(
      "<title>Brief note</title><nav>Navigation links</nav><main><p>A short but useful article.</p></main>",
      { url: "https://example.test/brief" },
    );
    const result = extractArticle(dom.window.document);

    expect(result.html).toContain("A short but useful article.");
    expect(result.html).not.toContain("Navigation links");
    expect(result.warnings).toHaveLength(1);
    dom.window.close();
  });

  it("extracts a generated 10,000-word article without losing sections or tables", () => {
    const dom = new JSDOM(createLargeFixture(10_000), {
      url: "https://example.test/large",
    });
    const result = extractArticle(dom.window.document);

    expect(result.stats.words).toBe(9_998);
    expect(result.sections).toHaveLength(5);
    expect(result.stats.tables).toBe(10);
    expect(result.stats.images).toBe(40);
    dom.window.close();
  });
});
