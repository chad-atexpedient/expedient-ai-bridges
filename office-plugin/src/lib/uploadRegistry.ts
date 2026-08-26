import type { AttachmentContext } from "./types";

export interface UploadedAsset {
  id: string;
  name: string;
  type: string;
  size: number;
  pixelWidth?: number;
  pixelHeight?: number;
  dataUrl: string;
  base64: string;
  kind?: "image" | "office-template";
  brandProfile?: AttachmentContext["brandProfile"];
}

const assets = new Map<string, UploadedAsset>();

function normalize(value: string) {
  return value.trim().toLowerCase();
}

export function registerUploadedAsset(asset: UploadedAsset) {
  assets.set(asset.id, asset);
  assets.set(normalize(asset.name), asset);
}

export function unregisterUploadedAsset(idOrName: string) {
  const asset = findUploadedAsset(idOrName);
  if (!asset) return;
  assets.delete(asset.id);
  assets.delete(normalize(asset.name));
}

export function clearUploadedAssets() {
  assets.clear();
}

export function findUploadedAsset(idOrName: unknown) {
  if (typeof idOrName !== "string" || !idOrName.trim()) return null;
  return assets.get(idOrName) ?? assets.get(normalize(idOrName)) ?? null;
}

export function listUploadedAssets() {
  const unique = new Map<string, UploadedAsset>();
  for (const asset of assets.values()) unique.set(asset.id, asset);
  return Array.from(unique.values()).map(({ id, name, type, size, pixelWidth, pixelHeight }) => ({ id, name, type, size, pixelWidth, pixelHeight }));
}

export async function imageAssetFromUrl(imageUrl: string) {
  const response = await fetch(`/api/image-asset?url=${encodeURIComponent(imageUrl)}`, { headers: { accept: "application/json" } });
  const json = await response.json().catch(() => null);
  if (!response.ok) throw new Error(json?.error?.message || "Image URL could not be fetched.");
  const asset = {
    id: imageUrl,
    name: json.name || imageUrl,
    type: json.type || "image/*",
    size: json.size || 0,
    pixelWidth: json.pixelWidth,
    pixelHeight: json.pixelHeight,
    dataUrl: json.dataUrl,
    base64: json.base64,
  } satisfies UploadedAsset;
  registerUploadedAsset(asset);
  return asset;
}

export function attachmentToUploadedAsset(attachment: AttachmentContext) {
  if (!attachment.assetKind || !attachment.dataUrl || !attachment.base64) return null;
  return {
    id: attachment.id,
    name: attachment.name,
    type: attachment.type,
    size: attachment.size,
    dataUrl: attachment.dataUrl,
    base64: attachment.base64,
    kind: attachment.assetKind,
    brandProfile: attachment.brandProfile,
  } satisfies UploadedAsset;
}
