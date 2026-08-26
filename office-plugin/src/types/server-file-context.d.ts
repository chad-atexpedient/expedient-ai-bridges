declare module "../../server/file-context.mjs" {
  export function handleFileContext(req: any, res: any): Promise<unknown>;
  export function extractUploadedFileContext(input: { name?: string; type?: string; base64: string; maxChars?: number }): {
    name: string;
    type: string;
    size: number;
    extractedText: string;
    clipped: boolean;
    strategy: string;
  };
}
