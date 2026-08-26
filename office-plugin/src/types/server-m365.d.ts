declare module "../../server/m365.mjs" {
  export function handleM365(req: any, res: any, url: URL): Promise<unknown>;
  export function extractOfficeText(buffer: Buffer, name?: string, mimeType?: string): string | null;
  export function rawFileFallback(buffer: Buffer, name?: string, mimeType?: string, maxChars?: number): string;
}
