import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";

const sizes = [10_000, 30_000, 100_000];
const words = ["article", "research", "example", "content", "reader", "document"];

export function createLargeFixture(targetWords) {
  if (!Number.isInteger(targetWords) || targetWords < 100) {
    throw new Error("Fixture size must be an integer of at least 100 words.");
  }
  const sectionCount = Math.ceil(targetWords / 2_000);
  const tableCount = Math.ceil(targetWords / 1_000);
  const imageCount = Math.ceil(targetWords / 250);
  let remainingWords = targetWords - 2 - sectionCount * 2 - tableCount * 6;
  const sections = [];

  for (let sectionIndex = 0; sectionIndex < sectionCount; sectionIndex += 1) {
    const sectionWords = Math.ceil(remainingWords / (sectionCount - sectionIndex));
    const paragraphs = [];
    for (let left = sectionWords; left > 0; left -= 120) {
      const count = Math.min(left, 120);
      paragraphs.push(
        `<p>${Array.from({ length: count }, (_, index) => words[index % words.length]).join(" ")}</p>`,
      );
    }
    remainingWords -= sectionWords;
    sections.push(
      `<section><h2>Section ${sectionIndex + 1}</h2>\n${paragraphs.join("\n")}</section>`,
    );
  }

  const tables = Array.from(
    { length: tableCount },
    () => "<table>\n<caption>Sample data</caption>\n<thead><tr><th>Field</th>\n<th>Value</th></tr></thead>\n<tbody><tr><td>Sample</td>\n<td>12</td></tr></tbody>\n</table>",
  ).join("\n");
  const images = Array.from(
    { length: imageCount },
    (_, index) => `<img src="https://example.test/image-${index + 1}.png" alt="">`,
  ).join("\n");
  const html = `<!doctype html><html lang="en"><head><title>Synthetic article</title></head><body><article><h1>Synthetic article</h1>\n${sections.join("\n")}\n${tables}\n${images}</article></body></html>`;
  const dom = new JSDOM(html);
  const measuredWords = dom.window.document.body.textContent?.trim().split(/\s+/).length ?? 0;
  dom.window.close();
  if (measuredWords !== targetWords) {
    throw new Error(`Expected ${targetWords} words, generated ${measuredWords}.`);
  }
  return html;
}

export async function generateLargeFixtures(
  outputDirectory = resolve(dirname(fileURLToPath(import.meta.url)), ".generated"),
) {
  await mkdir(outputDirectory, { recursive: true });
  for (const targetWords of sizes) {
    await writeFile(
      resolve(outputDirectory, `${targetWords}-words.html`),
      createLargeFixture(targetWords),
      "utf8",
    );
    console.log(
      `${targetWords.toLocaleString()} words, ${Math.ceil(targetWords / 250)} images, ${Math.ceil(targetWords / 1_000)} tables`,
    );
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await generateLargeFixtures();
}
