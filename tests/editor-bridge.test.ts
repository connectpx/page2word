import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  documentAssetKey,
  documentMetaKey,
  documentSectionKey,
} from "../src/shared/storage-keys";

const documentId = "document-bridge-123";
const assetId = "asset-bridge-123";
const base64 = "AQID";
const origin = "http://localhost:5173";

type BridgeEvent = {
  source: unknown;
  origin: string;
  data: unknown;
};

describe("editor asset bridge", () => {
  let listener: ((event: BridgeEvent) => void) | undefined;
  let postedMessages: Record<string, unknown>[];
  let postMessageCalls: unknown[][];
  let storageChangeListener:
    | ((
        changes: Record<string, { newValue?: unknown }>,
        areaName: string,
      ) => void)
    | undefined;
  let storedValues: Map<string, unknown>;
  let windowMock: {
    addEventListener: (
      type: string,
      callback: (event: BridgeEvent) => void,
    ) => void;
    postMessage: (...args: unknown[]) => void;
  };

  beforeEach(async () => {
    vi.resetModules();
    vi.unstubAllGlobals();
    listener = undefined;
    storageChangeListener = undefined;
    postedMessages = [];
    postMessageCalls = [];
    storedValues = new Map<string, unknown>([
      [
        documentMetaKey(documentId),
        {
          id: documentId,
          sectionCount: 1,
          sectionsAvailable: true,
          sections: [{ id: "section-1", title: "Article", words: 1 }],
          assets: [
            {
              id: assetId,
              mime: "image/png",
              width: 1,
              height: 1,
              bytes: 3,
            },
          ],
        },
      ],
      [
        documentSectionKey(documentId, 0),
        {
          id: "section-1",
          title: "Article",
          words: 1,
          html: "<p>Article</p>",
        },
      ],
      [documentAssetKey(documentId, assetId), base64],
    ]);
    windowMock = {
      addEventListener: (_type, callback) => {
        listener = callback;
      },
      postMessage: (...args) => {
        postMessageCalls.push(args);
        const message = args[0];
        if (isRecord(message)) {
          postedMessages.push(message);
        }
      },
    };
    vi.stubGlobal("window", windowMock);
    vi.stubGlobal("location", {
      hash: `#doc=${documentId}`,
      origin,
    });
    vi.stubGlobal("chrome", {
      storage: {
        onChanged: {
          addListener: (callback: typeof storageChangeListener) => {
            storageChangeListener = callback;
          },
        },
        local: {
          get: vi.fn(async (keys: string | string[]) => {
            const requested = Array.isArray(keys) ? keys : [keys];
            return Object.fromEntries(
              requested
                .filter((key) => storedValues.has(key))
                .map((key) => [key, storedValues.get(key)]),
            );
          }),
        },
      },
      runtime: { sendMessage: vi.fn() },
    });
    await import("../src/extension/bridge/editor-bridge");
  });

  it("streams captured image data as a base64 message before document end", async () => {
    if (!listener) {
      throw new Error("The editor bridge did not register its listener.");
    }

    listener({
      source: windowMock,
      origin,
      data: { source: "page2word-editor", type: "READY" },
    });
    await vi.waitFor(() =>
      expect(postedMessages.some((message) => message.type === "DOC_END")).toBe(
        true,
      ),
    );

    expect(postedMessages.map((message) => message.type)).toEqual([
      "DOC_BEGIN",
      "SECTION",
      "ASSET",
      "DOC_END",
    ]);
    expect(postedMessages[2]).toMatchObject({
      asset: { id: assetId, base64 },
      type: "ASSET",
    });
    expect(postMessageCalls[2]).toHaveLength(2);
  });

  it("reports a missing image but still completes the text handoff", async () => {
    if (!listener) {
      throw new Error("The editor bridge did not register its listener.");
    }
    storedValues.delete(documentAssetKey(documentId, assetId));

    listener({
      source: windowMock,
      origin,
      data: { source: "page2word-editor", type: "READY" },
    });
    await vi.waitFor(() =>
      expect(postedMessages.some((message) => message.type === "DOC_END")).toBe(
        true,
      ),
    );

    expect(postedMessages).toContainEqual(
      expect.objectContaining({
        assetId,
        type: "ASSET_ERROR",
      }),
    );
    expect(postedMessages.at(-1)?.type).toBe("DOC_END");
    expect(postedMessages.some((message) => message.type === "ERROR")).toBe(
      false,
    );
  });

  it("forwards locally changed formatting settings to the editor", () => {
    storageChangeListener?.(
      { docxSettings: { newValue: { fontSizePt: 15 } } },
      "local",
    );

    expect(postedMessages.at(-1)).toMatchObject({
      source: "page2word",
      type: "SETTINGS",
      settings: { fontSizePt: 15 },
    });
  });
});

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
