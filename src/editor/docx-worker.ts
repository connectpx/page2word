import { createDocxDocument, packDocx } from "../shared/html-to-docx/document";
import type {
  DocxBlock,
  DocxImageAsset,
  DocxMetadata,
  DocxSettings,
} from "../shared/html-to-docx/types";

interface ExportRequest {
  type: "EXPORT";
  blocks: DocxBlock[];
  assets: DocxImageAsset[];
  metadata: DocxMetadata;
  settings: DocxSettings;
}

self.addEventListener("message", (event: MessageEvent<ExportRequest>) => {
  if (event.data?.type !== "EXPORT") {
    return;
  }
  void generate(event.data);
});

async function generate(request: ExportRequest): Promise<void> {
  try {
    self.postMessage({ type: "PROGRESS", stage: "building" });
    const document = createDocxDocument(
      request.blocks,
      request.metadata,
      request.settings,
      new Map(request.assets.map((asset) => [asset.id, asset])),
    );
    const blob = await packDocx(document);
    self.postMessage({ type: "DONE", blob });
  } catch (error) {
    self.postMessage({
      type: "ERROR",
      error:
        error instanceof Error
          ? error.message
          : "The Word document could not be generated.",
    });
  }
}
