const button = document.querySelector<HTMLButtonElement>("#extract");
const browserEditorButton =
  document.querySelector<HTMLButtonElement>("#extract-browser");
const openEditorButton =
  document.querySelector<HTMLButtonElement>("#open-editor");
const statusElement = document.querySelector<HTMLElement>("#status");
const serverStatusElement =
  document.querySelector<HTMLElement>("#server-status");
const serverStatusTextElement = document.querySelector<HTMLElement>(
  "#server-status-text",
);
const serverCheckButton =
  document.querySelector<HTMLButtonElement>("#check-server");
const serverHelpElement = document.querySelector<HTMLElement>("#server-help");
const BROWSER_EDITOR_HEALTH_URL = "http://localhost:5173/health";

button?.addEventListener("click", () => void extractArticle("built-in"));
browserEditorButton?.addEventListener(
  "click",
  () => void extractArticle("browser"),
);
serverCheckButton?.addEventListener(
  "click",
  () => void checkBrowserEditorServer(),
);

void checkBrowserEditorServer();

openEditorButton?.addEventListener("click", async () => {
  if (!statusElement || !openEditorButton) {
    return;
  }
  openEditorButton.disabled = true;
  statusElement.textContent = "Opening the editor…";
  try {
    await chrome.tabs.create({
      url: chrome.runtime.getURL("editor/index.html"),
    });
    statusElement.textContent = "Editor opened.";
  } catch (error) {
    statusElement.textContent =
      error instanceof Error ? error.message : "Could not open the editor.";
  } finally {
    openEditorButton.disabled = false;
  }
});

async function extractArticle(
  editorMode: "built-in" | "browser",
): Promise<void> {
  const activeButton = editorMode === "browser" ? browserEditorButton : button;
  if (!activeButton || !statusElement) {
    return;
  }
  activeButton.disabled = true;
  statusElement.textContent =
    editorMode === "browser"
      ? "Checking editor server and reading this page…"
      : "Reading this page…";
  try {
    const [tab] = await chrome.tabs.query({
      active: true,
      currentWindow: true,
    });
    if (!tab?.id) {
      throw new Error("No active tab is available.");
    }
    const response = await chrome.runtime.sendMessage({
      type: "EXTRACT",
      tabId: tab.id,
      editorMode,
    });
    if (!response?.ok) {
      throw new Error(response?.error ?? "Article extraction failed.");
    }
    const destination =
      response.openedIn === "browser" ? "browser editor" : "built-in editor";
    statusElement.textContent = `Opened "${response.title}" in the ${destination} (${response.words.toLocaleString()} words).`;
  } catch (error) {
    statusElement.textContent =
      error instanceof Error && error.message.includes("Cannot access")
        ? "This page can't be read. Try a regular web page instead of a browser settings page or PDF."
        : error instanceof Error
          ? error.message
          : "Article extraction failed.";
    if (
      editorMode === "browser" &&
      error instanceof Error &&
      error.message.includes("Node editor server is not running")
    ) {
      if (serverStatusElement) {
        setBrowserEditorStatus("offline", error.message);
      }
    }
  } finally {
    activeButton.disabled = false;
  }
}

async function checkBrowserEditorServer(): Promise<void> {
  if (!serverStatusElement || !serverStatusTextElement || !serverCheckButton) {
    return;
  }

  serverCheckButton.disabled = true;
  setBrowserEditorStatus("checking", "Checking server…");
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 1_500);

  try {
    const response = await fetch(BROWSER_EDITOR_HEALTH_URL, {
      cache: "no-store",
      signal: controller.signal,
    });
    const health: unknown = await response.json();
    if (
      !response.ok ||
      typeof health !== "object" ||
      health === null ||
      !("service" in health) ||
      health.service !== "page2word-editor"
    ) {
      throw new Error("Editor server health check failed.");
    }
    setBrowserEditorStatus("online", "Server is running");
  } catch {
    setBrowserEditorStatus("offline", "Server is not running");
  } finally {
    window.clearTimeout(timeout);
    serverCheckButton.disabled = false;
  }
}

function setBrowserEditorStatus(
  state: "checking" | "online" | "offline",
  message: string,
): void {
  if (!serverStatusElement || !serverStatusTextElement) {
    return;
  }
  serverStatusElement.dataset.state = state;
  serverStatusTextElement.textContent = message;
  if (serverHelpElement) {
    serverHelpElement.hidden = state !== "offline";
  }
}

export {};
