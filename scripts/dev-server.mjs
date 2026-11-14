import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, isAbsolute, relative, resolve, sep } from "node:path";

const root = resolve(process.cwd(), "dist", "editor");
const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
};

createServer(async (request, response) => {
  try {
    if (request.method !== "GET") {
      response.writeHead(405, { Allow: "GET" });
      response.end("Method not allowed");
      return;
    }

    const requestPath = decodeURIComponent(
      new URL(request.url, "http://localhost").pathname,
    );
    if (requestPath === "/health") {
      response.writeHead(200, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(JSON.stringify({ service: "page2word-editor" }));
      return;
    }

    const relativePath =
      requestPath === "/" ? "index.html" : requestPath.slice(1);
    const filePath = resolve(root, relativePath);
    const pathFromRoot = relative(root, filePath);
    if (
      pathFromRoot === ".." ||
      pathFromRoot.startsWith(`..${sep}`) ||
      isAbsolute(pathFromRoot)
    ) {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
      return;
    }

    const content = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type":
        contentTypes[extname(filePath)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
    });
    response.end(content);
  } catch {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
  }
}).listen(5173, "127.0.0.1", () => {
  console.log("Page2Word editor: http://localhost:5173");
});
