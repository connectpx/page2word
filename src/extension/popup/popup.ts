const button = document.querySelector<HTMLButtonElement>("#extract");
const browserEditorButton =
  document.querySelector<HTMLButtonElement>("#extract-browser");
const openEditorButton =
  document.querySelector<HTMLButtonElement>("#open-editor");
const statusElement = document.querySelector<HTMLElement>("#status");
const serverStatusElement =
  document.querySelector<HTMLElement>("#server-status");

button?.addEventListener("click", () => void extractArticle("built-in"));
browserEditorButton?.addEventListener(
  "click",
  () => void extractArticle("browser"),
);

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
        serverStatusElement.textContent = error.message;
        serverStatusElement.classList.add("is-error");
        serverStatusElement.classList.remove("is-ready");
      }
    }
  } finally {
    activeButton.disabled = false;
  }
}

export {};
