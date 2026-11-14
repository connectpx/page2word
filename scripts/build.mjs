import { cp, mkdir, rm } from "node:fs/promises";
import { resolve } from "node:path";
import { build } from "esbuild";

const root = process.cwd();
const extensionOut = resolve(root, "dist/extension");
const editorOut = resolve(root, "dist/editor");
const bundledEditorOut = resolve(extensionOut, "editor");

await rm(resolve(root, "dist"), { recursive: true, force: true });
await Promise.all([
  mkdir(extensionOut, { recursive: true }),
  mkdir(editorOut, { recursive: true }),
  mkdir(resolve(extensionOut, "popup"), { recursive: true }),
  mkdir(resolve(extensionOut, "options"), { recursive: true }),
  mkdir(resolve(extensionOut, "background"), { recursive: true }),
  mkdir(resolve(extensionOut, "bridge"), { recursive: true }),
  mkdir(bundledEditorOut, { recursive: true }),
  mkdir(resolve(extensionOut, "extractor"), { recursive: true }),
]);

await Promise.all([
  cp(
    resolve(root, "src/extension/manifest.json"),
    resolve(extensionOut, "manifest.json"),
  ),
  cp(
    resolve(root, "src/extension/popup/popup.html"),
    resolve(extensionOut, "popup/popup.html"),
  ),
  cp(
    resolve(root, "src/extension/popup/popup.css"),
    resolve(extensionOut, "popup/popup.css"),
  ),
  cp(
    resolve(root, "src/extension/options/options.html"),
    resolve(extensionOut, "options/options.html"),
  ),
  cp(
    resolve(root, "src/extension/options/options.css"),
    resolve(extensionOut, "options/options.css"),
  ),
  cp(resolve(root, "src/editor/index.html"), resolve(editorOut, "index.html")),
  cp(resolve(root, "src/editor/editor.css"), resolve(editorOut, "editor.css")),
  cp(
    resolve(root, "src/editor/index.html"),
    resolve(bundledEditorOut, "index.html"),
  ),
  cp(
    resolve(root, "src/editor/editor.css"),
    resolve(bundledEditorOut, "editor.css"),
  ),
]);

await Promise.all([
  build({
    entryPoints: [resolve(root, "src/extension/background/service-worker.ts")],
    bundle: true,
    outfile: resolve(extensionOut, "background/service-worker.js"),
    platform: "browser",
    target: "chrome120",
    format: "esm",
  }),
  build({
    entryPoints: [
      resolve(root, "src/extension/popup/popup.ts"),
      resolve(root, "src/extension/options/options.ts"),
      resolve(root, "src/extension/bridge/editor-bridge.ts"),
      resolve(root, "src/extension/extractor/extract.ts"),
    ],
    bundle: true,
    outbase: resolve(root, "src/extension"),
    outdir: extensionOut,
    platform: "browser",
    target: "chrome120",
    format: "iife",
  }),
  build({
    entryPoints: [resolve(root, "src/editor/editor.ts")],
    bundle: true,
    outfile: resolve(bundledEditorOut, "editor.js"),
    platform: "browser",
    target: "chrome120",
    format: "iife",
  }),
  build({
    entryPoints: [resolve(root, "src/editor/docx-worker.ts")],
    bundle: true,
    outfile: resolve(bundledEditorOut, "docx-worker.js"),
    platform: "browser",
    target: "chrome120",
    format: "iife",
  }),
]);

await Promise.all(
  ["editor.js", "docx-worker.js"].map((file) =>
    cp(resolve(bundledEditorOut, file), resolve(editorOut, file)),
  ),
);
