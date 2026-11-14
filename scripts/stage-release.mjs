import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = process.cwd();
const packageJson = JSON.parse(
  await readFile(resolve(root, "package.json"), "utf8"),
);
const version = packageJson.version;
const tag = process.env.GITHUB_REF_NAME;
const manifest = JSON.parse(
  await readFile(resolve(root, "src", "extension", "manifest.json"), "utf8"),
);

if (
  typeof version !== "string" ||
  !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(version)
) {
  throw new Error("package.json must contain a semantic version.");
}
if (manifest.version !== version) {
  throw new Error(
    `Extension manifest version ${manifest.version} does not match package version ${version}.`,
  );
}
if (tag && tag !== `v${version}`) {
  throw new Error(
    `Release tag ${tag} does not match package version ${version}.`,
  );
}

const releaseName = `Page2Word-v${version}`;
const releaseRoot = resolve(root, "dist", "release", releaseName);
const extensionPath = resolve(root, "dist", "extension");
const editorPath = resolve(root, "dist", "editor");

await Promise.all([
  readFile(resolve(extensionPath, "manifest.json")),
  readFile(resolve(editorPath, "index.html")),
]);
await rm(releaseRoot, { recursive: true, force: true });
await Promise.all([
  mkdir(resolve(releaseRoot, "dist"), { recursive: true }),
  mkdir(resolve(releaseRoot, "scripts"), { recursive: true }),
]);
await Promise.all([
  cp(extensionPath, resolve(releaseRoot, "dist", "extension"), {
    recursive: true,
  }),
  cp(editorPath, resolve(releaseRoot, "dist", "editor"), {
    recursive: true,
  }),
  cp(
    resolve(root, "scripts", "dev-server.mjs"),
    resolve(releaseRoot, "scripts", "dev-server.mjs"),
  ),
]);
await writeFile(
  resolve(releaseRoot, "README.txt"),
  `Page2Word ${version}

Install the extension:
1. Open chrome://extensions and enable Developer mode.
2. Select Load unpacked and choose dist\\extension.

The built-in editor works without Node.js.

To use writing-assistant extensions:
1. Install Node.js 20 or newer.
2. From this folder, run: node scripts\\dev-server.mjs
3. Keep the terminal open and choose Extract with Writing Assistants.

The editor server serves the included dist\\editor files at http://localhost:5173.
No npm install or rebuild is required.
`,
  "utf8",
);

console.log(`Release package staged at ${releaseRoot}`);
