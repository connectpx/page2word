import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const manifest = JSON.parse(
  readFileSync(
    resolve(
      dirname(fileURLToPath(import.meta.url)),
      "../src/extension/manifest.json",
    ),
    "utf8",
  ),
) as {
  manifest_version: number;
  background: { service_worker: string; type: string };
  action: { default_popup: string };
  content_scripts: { js: string[]; matches: string[] }[];
};

describe("extension manifest", () => {
  it("uses MV3 and points to files emitted by the build", () => {
    expect(manifest.manifest_version).toBe(3);
    expect(manifest.background).toEqual({
      service_worker: "background/service-worker.js",
      type: "module",
    });
    expect(manifest.action.default_popup).toBe("popup/popup.html");
    expect(manifest.content_scripts[0].js).toEqual(["bridge/editor-bridge.js"]);
    expect(manifest.content_scripts[0].matches).toContain(
      "http://localhost:5173/*",
    );
  });
});
