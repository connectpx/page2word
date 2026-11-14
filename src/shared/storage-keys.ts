export function documentMetaKey(id: string): string {
  return `page2word:document:${id}:meta`;
}

export function documentSectionKey(id: string, index: number): string {
  return `page2word:document:${id}:section:${index}`;
}

export function documentAssetKey(id: string, assetId: string): string {
  return `page2word:document:${id}:asset:${assetId}`;
}
