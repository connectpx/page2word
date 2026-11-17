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
    fetchMock = vi.fn(
      async () =>
        new Response(JSON.stringify({ service: "page2word-editor" }), {
          status: 200,
        }),
    );
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

  it("shows server status without waiting for the status request", async () => {
    let resolveFetch: ((response: Response) => void) | undefined;
    fetchMock.mockImplementation(
      () =>
        new Promise<Response>((resolve) => {
          resolveFetch = resolve;
        }),
    );
    await import("../src/extension/popup/popup");

    expect(fetchMock).toHaveBeenCalledWith(
      "http://localhost:5173/health",
      expect.objectContaining({ cache: "no-store" }),
    );
    expect(
      dom.window.document.querySelector<HTMLButtonElement>("#extract-browser")
        ?.disabled,
    ).toBe(false);
    resolveFetch?.(
      new Response(JSON.stringify({ service: "page2word-editor" }), {
        status: 200,
      }),
    );
    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelector<HTMLElement>("#server-status")
          ?.dataset.state,
      ).toBe("online");
    });
  });

  it("shows an offline indicator and can recheck when the server starts", async () => {
    fetchMock.mockRejectedValueOnce(new Error("Connection refused"));
    await import("../src/extension/popup/popup");

    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelector<HTMLElement>("#server-status")
          ?.dataset.state,
      ).toBe("offline");
    });
    expect(
      dom.window.document.querySelector<HTMLElement>("#server-help")?.hidden,
    ).toBe(false);

    dom.window.document
      .querySelector<HTMLButtonElement>("#check-server")
      ?.click();
    await vi.waitFor(() => {
      expect(
        dom.window.document.querySelector<HTMLElement>("#server-status")
          ?.dataset.state,
      ).toBe("online");
    });
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
        'The Node editor server is not running. Double-click "StartServer.bat" in the Page2Word folder, then try again.',
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
    expect(fetchMock).toHaveBeenCalled();
    expect(
      dom.window.document.querySelector<HTMLButtonElement>("#extract-browser")
        ?.disabled,
    ).toBe(false);
  });
});
