import type {
  ContentSection,
  ExtractedAsset,
  StoredDocument,
} from "../shared/types";

export interface EditorDocumentRecord extends StoredDocument {
  updatedAt: string;
}

export interface EditorSectionRecord {
  documentId: string;
  id: string;
  title: string;
  words: number;
  html: string;
  originalHtml?: string;
  done: boolean;
  hidden?: boolean;
}

export interface EditorAssetRecord extends ExtractedAsset {
  documentId: string;
  blob: Blob;
}

let databasePromise: Promise<IDBDatabase> | undefined;

function openDatabase(): Promise<IDBDatabase> {
  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = indexedDB.open("page2word-editor", 3);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains("documents")) {
          database.createObjectStore("documents", { keyPath: "id" });
        }
        if (!database.objectStoreNames.contains("sections")) {
          database.createObjectStore("sections", {
            keyPath: ["documentId", "id"],
          });
        }
        if (!database.objectStoreNames.contains("assets")) {
          database.createObjectStore("assets", {
            keyPath: ["documentId", "id"],
          });
        }
      };
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onerror = () => {
        databasePromise = undefined;
        reject(
          request.error ?? new Error("Could not open local document storage."),
        );
      };
      request.onblocked = () => {
        databasePromise = undefined;
        reject(
          new Error("Local document storage is blocked by another editor tab."),
        );
      };
    });
  }
  return databasePromise;
}

function transactionComplete(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () =>
      reject(transaction.error ?? new Error("Local save was aborted."));
    transaction.onerror = () =>
      reject(transaction.error ?? new Error("Local save failed."));
  });
}

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () =>
      reject(request.error ?? new Error("Local storage request failed."));
  });
}

export async function loadEditorDocument(id: string): Promise<{
  document?: EditorDocumentRecord;
  sections: EditorSectionRecord[];
  assets: EditorAssetRecord[];
}> {
  const database = await openDatabase();
  const transaction = database.transaction(
    ["documents", "sections", "assets"],
    "readonly",
  );
  const [document, sections, assets] = await Promise.all([
    requestResult(transaction.objectStore("documents").get(id)) as Promise<
      EditorDocumentRecord | undefined
    >,
    requestResult(transaction.objectStore("sections").getAll()) as Promise<
      EditorSectionRecord[]
    >,
    requestResult(transaction.objectStore("assets").getAll()) as Promise<
      EditorAssetRecord[]
    >,
  ]);
  return {
    document,
    sections: sections.filter((section) => section.documentId === id),
    assets: assets.filter((asset) => asset.documentId === id),
  };
}

export async function saveEditorDocument(
  document: EditorDocumentRecord,
  sections: EditorSectionRecord[],
  assets: EditorAssetRecord[] = [],
): Promise<void> {
  const database = await openDatabase();
  const transaction = database.transaction(
    ["documents", "sections", "assets"],
    "readwrite",
  );
  transaction.objectStore("documents").put(document);
  const sectionStore = transaction.objectStore("sections");
  for (const section of sections) {
    sectionStore.put(section);
  }
  const assetStore = transaction.objectStore("assets");
  const keepAssetIds = new Set(assets.map((asset) => asset.id));
  const existingAssets = assetStore.getAll();
  existingAssets.onsuccess = () => {
    for (const asset of existingAssets.result as EditorAssetRecord[]) {
      if (asset.documentId === document.id && !keepAssetIds.has(asset.id)) {
        assetStore.delete([asset.documentId, asset.id]);
      }
    }
  };
  for (const asset of assets) {
    assetStore.put(asset);
  }
  await transactionComplete(transaction);
}

export async function saveEditorSections(
  documentId: string,
  sections: EditorSectionRecord[],
): Promise<void> {
  const database = await openDatabase();
  const transaction = database.transaction(
    ["documents", "sections"],
    "readwrite",
  );
  const documentStore = transaction.objectStore("documents");
  const document = (await requestResult(documentStore.get(documentId))) as
    EditorDocumentRecord | undefined;
  if (!document) {
    transaction.abort();
    throw new Error("The edited document is missing from local storage.");
  }
  const updatedSections = new Map(
    sections.map((section) => [section.id, section]),
  );
  document.sections = document.sections.map((section) => {
    const updated = updatedSections.get(section.id);
    return updated ? { ...section, words: updated.words } : section;
  });
  document.stats.words = document.sections.reduce(
    (total, section) => total + section.words,
    0,
  );
  document.updatedAt = new Date().toISOString();
  documentStore.put(document);
  const sectionStore = transaction.objectStore("sections");
  for (const section of sections) {
    sectionStore.put(section);
  }
  await transactionComplete(transaction);
}

export async function deleteEditorDocument(id: string): Promise<void> {
  const database = await openDatabase();
  const transaction = database.transaction(
    ["documents", "sections", "assets"],
    "readwrite",
  );
  transaction.objectStore("documents").delete(id);
  const sectionStore = transaction.objectStore("sections");
  const sections = (await requestResult(
    sectionStore.getAll(),
  )) as EditorSectionRecord[];
  for (const section of sections) {
    if (section.documentId === id) {
      sectionStore.delete([section.documentId, section.id]);
    }
  }
  const assetStore = transaction.objectStore("assets");
  const assets = (await requestResult(
    assetStore.getAll(),
  )) as EditorAssetRecord[];
  for (const asset of assets) {
    if (asset.documentId === id) {
      assetStore.delete([asset.documentId, asset.id]);
    }
  }
  await transactionComplete(transaction);
}

export function toEditorSection(
  documentId: string,
  section: ContentSection,
  html = section.html,
  done = false,
): EditorSectionRecord {
  return {
    documentId,
    id: section.id,
    title: section.title,
    words: section.words,
    html,
    done,
  };
}
