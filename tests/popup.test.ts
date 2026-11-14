import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

describe("popup writing-assistant mode", () => {
  let dom: JSDOM;
  let fetchMock: ReturnType<typeof vi.fn>;
  let sendMessage: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    vi.resetModules();
    vi.unstubAllGlobals();
    const popupHtml = readFileSync(
      resolve(
        dirname(fileURLToPath(import.meta.url)),
        "../src/extension/popup/popup.html",
      ),
      "utf8",
    );
    dom = new JSDOM(popupHtml, {
      url: "chrome-extension://test-id/popup/popup.html",
    });
    fetchMock = vi.fn();
    sendMessage = vi.fn(async () => ({
      ok: true,
      title: "Test article",
      words: 10,
      openedIn: "browser",
    }));
    vi.stubGlobal("window", dom.window);
    vi.stubGlobal("document", dom.window.document);
    vi.stubGlobal("fetch", fetchMock);
    vi.stubGlobal("chrome", {
      tabs: {
        create: vi.fn(async () => undefined),
        query: vi.fn(async () => [{ id: 12 }]),
      },
      runtime: {
        getURL: (path: string) => `chrome-extension://test-id/${path}`,
        sendMessage,
      },
    });
  });

  afterEach(() => {
    dom.window.close();
  });

  it("opens without waiting for a server status request", async () => {
    await import("../src/extension/popup/popup");

    expect(fetchMock).not.toHaveBeenCalled();
    expect(
      dom.window.document.querySelector<HTMLButtonElement>("#extract-browser")
        ?.disabled,
    ).toBe(false);
  });

  it("requests browser editing only when the user chooses the mode", async () => {
    await import("../src/extension/popup/popup");
    const extractButton =
      dom.window.document.querySelector<HTMLButtonElement>("#extract-browser");

    extractButton?.click();
    await vi.waitFor(() => expect(sendMessage).toHaveBeenCalledOnce());

    expect(sendMessage).toHaveBeenCalledWith({
      type: "EXTRACT",
      tabId: 12,
      editorMode: "browser",
    });
  });

  it("shows how to start the server after a failed browser-mode request", async () => {
    sendMessage.mockResolvedValue({
      ok: false,
      error:
        'The Node editor server is not running. From the project folder run "node scripts\\dev-server.mjs", then try again.',
    });
    await import("../src/extension/popup/popup");
    dom.window.document
      .querySelector<HTMLButtonElement>("#extract-browser")
      ?.click();

    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelector<HTMLElement>("#server-status")
          ?.textContent,
      ).toContain("Node editor server is not running");
    });
    expect(fetchMock).not.toHaveBeenCalled();
    expect(
      dom.window.document.querySelector<HTMLButtonElement>("#extract-browser")
        ?.disabled,
    ).toBe(false);
  });
});
