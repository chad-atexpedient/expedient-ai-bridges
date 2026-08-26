import { extractOfficeText, extractPowerPointBrandProfile, officePackageHasPart, rawFileFallback } from "./m365.mjs";

// Read at call time, not module load. A module-load constant made the limit
// silently depend on import order relative to environment configuration.
function maxUploadBytes() {
  const configured = Number(process.env.FILE_CONTEXT_MAX_UPLOAD_BYTES);
  return Number.isFinite(configured) && configured > 0 ? configured : 25 * 1024 * 1024;
}
const DEFAULT_MAX_CHARS = 12000;

function sendJson(res, status, body) {
  res.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

class PayloadTooLargeError extends Error {}

async function readJson(req, limitBytes = Math.ceil(maxUploadBytes() * 1.4) + 4096) {
  const chunks = [];
  let total = 0;
  // Enforce the bound while streaming. Buffering first and checking after
  // meant an oversized upload was fully resident in memory before rejection.
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > limitBytes) {
      // Destroy the socket rather than just throwing. Abandoning the stream
      // without destroying it left the client free to keep sending, so the
      // size limit stopped memory growth but not the flood itself.
      req.destroy();
      throw new PayloadTooLargeError(`File context payload is too large. Limit is ${maxUploadBytes()} bytes before base64 encoding.`);
    }
    chunks.push(buffer);
  }
  const body = chunks.length ? Buffer.concat(chunks) : undefined;
  if (!body?.length) return null;
  return JSON.parse(body.toString("utf8"));
}

function textLikeMime(mimeType = "", name = "") {
  return mimeType.startsWith("text/") || /\.(txt|csv|tsv|json|md|markdown|log|xml|html|css|js|ts|tsx|jsx|py|ps1|sql|yaml|yml)$/i.test(name) || ["application/json", "application/xml"].includes(mimeType);
}

function officePackageType(name = "", type = "") {
  const lower = `${name} ${type}`.toLowerCase();
  if (lower.includes("wordprocessingml.document") || /\.docx$/i.test(name)) return "docx";
  if (lower.includes("presentationml.presentation") || /\.pptx$/i.test(name)) return "pptx";
  if (lower.includes("spreadsheetml.sheet") || /\.xlsx$/i.test(name)) return "xlsx";
  return null;
}

function declaredOfficeExtension(name = "") {
  return String(name).toLowerCase().match(/\.(docx|pptx|xlsx)$/i)?.[1] || null;
}

function validateDeclaredFileType(name = "", type = "", packageType = null) {
  const extension = declaredOfficeExtension(name);
  if (!extension || !packageType || extension === packageType) return;
  throw new Error(`The uploaded file name and MIME type disagree: .${extension} was declared as ${packageType.toUpperCase()}.`);
}

function validateOfficePackageSignature(buffer, packageType) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 4 || !buffer.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
    throw new Error(`The uploaded ${packageType.toUpperCase()} file does not have a valid Office package signature.`);
  }
  const requiredPart = packageType === "docx" ? "word/document.xml" : packageType === "pptx" ? "ppt/presentation.xml" : "xl/workbook.xml";
  if (!officePackageHasPart(buffer, "[Content_Types].xml")) throw new Error(`The uploaded ${packageType.toUpperCase()} file is missing its package content types.`);
  if (!officePackageHasPart(buffer, requiredPart)) throw new Error(`The uploaded ${packageType.toUpperCase()} file is missing ${requiredPart}.`);
}

export function extractUploadedFileContext({ name = "uploaded-file", type = "", base64 = "", maxChars = DEFAULT_MAX_CHARS }) {
  if (!base64 || typeof base64 !== "string") throw new Error("Missing base64 file content.");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(base64) || base64.length % 4 === 1) throw new Error("Invalid base64 file content.");
  const buffer = Buffer.from(base64, "base64");
  if (buffer.length > maxUploadBytes()) throw new Error(`File ${name} is too large for context extraction. Limit is ${(maxUploadBytes() / 1024 / 1024).toFixed(0)} MB.`);

  const isTextLike = textLikeMime(type, name);
  const packageType = officePackageType(name, type);
  validateDeclaredFileType(name, type, packageType);
  if (packageType) validateOfficePackageSignature(buffer, packageType);
  const officeText = isTextLike ? null : extractOfficeText(buffer, name, type);
  const brandProfile = packageType === "pptx" ? extractPowerPointBrandProfile(buffer) : undefined;
  let text = isTextLike ? buffer.toString("utf8") : officeText;
  if (!text) text = rawFileFallback(buffer, name, type, maxChars);
  const limit = Math.max(1000, Number(maxChars) || DEFAULT_MAX_CHARS);

  return {
    name,
    type: type || "application/octet-stream",
    size: buffer.length,
    extractedText: text.slice(0, limit),
    clipped: text.length > limit,
    strategy: isTextLike ? "text" : officeText ? "office-open-xml" : "raw-fallback",
    ...(brandProfile ? { brandProfile } : {}),
  };
}

export async function handleFileContext(req, res) {
  if (req.method !== "POST") return sendJson(res, 405, { error: { message: "Method not allowed" } });
  try {
    const body = await readJson(req);
    if (!body) return sendJson(res, 400, { error: { message: "Missing file context payload" } });
    return sendJson(res, 200, extractUploadedFileContext(body));
  } catch (error) {
    const status = error instanceof PayloadTooLargeError ? 413 : 400;
    if (res.writableEnded || res.destroyed) return undefined;
    return sendJson(res, status, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}
