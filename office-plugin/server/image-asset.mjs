import { assertSafeOutboundUrl, safeFilename } from "./security.mjs";
const MAX_IMAGE_BYTES = 8_000_000;
const MAX_IMAGE_PIXELS = Number(process.env.IMAGE_MAX_PIXELS || 40_000_000);
const IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"]);

function sendJson(res, status, body) {
  res.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

function extensionFromType(type = "") {
  if (type.includes("png")) return "png";
  if (type.includes("jpeg") || type.includes("jpg")) return "jpg";
  if (type.includes("gif")) return "gif";
  if (type.includes("webp")) return "webp";
  if (type.includes("bmp")) return "bmp";
  return "bin";
}

function readUInt24BE(buffer, offset) {
  return (buffer[offset] << 16) + (buffer[offset + 1] << 8) + buffer[offset + 2];
}

function imageDimensions(bytes, type = "") {
  if (!Buffer.isBuffer(bytes) || bytes.length < 10) return {};
  if (type.includes("png") && bytes.length >= 24 && bytes.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    return { pixelWidth: bytes.readUInt32BE(16), pixelHeight: bytes.readUInt32BE(20) };
  }
  if (type.includes("gif") && bytes.length >= 10 && bytes.slice(0, 3).toString("ascii") === "GIF") {
    return { pixelWidth: bytes.readUInt16LE(6), pixelHeight: bytes.readUInt16LE(8) };
  }
  if (type.includes("bmp") && bytes.length >= 26 && bytes.slice(0, 2).toString("ascii") === "BM") {
    return { pixelWidth: Math.abs(bytes.readInt32LE(18)), pixelHeight: Math.abs(bytes.readInt32LE(22)) };
  }
  if (type.includes("webp") && bytes.length >= 30 && bytes.slice(0, 4).toString("ascii") === "RIFF" && bytes.slice(8, 12).toString("ascii") === "WEBP") {
    const chunk = bytes.slice(12, 16).toString("ascii");
    if (chunk === "VP8 " && bytes.length >= 30) return { pixelWidth: bytes.readUInt16LE(26) & 0x3fff, pixelHeight: bytes.readUInt16LE(28) & 0x3fff };
    if (chunk === "VP8L" && bytes.length >= 25) {
      const b0 = bytes[21];
      const b1 = bytes[22];
      const b2 = bytes[23];
      const b3 = bytes[24];
      return { pixelWidth: 1 + (((b1 & 0x3f) << 8) | b0), pixelHeight: 1 + (((b3 & 0x0f) << 10) | (b2 << 2) | ((b1 & 0xc0) >> 6)) };
    }
    if (chunk === "VP8X" && bytes.length >= 30) return { pixelWidth: 1 + readUInt24BE(Buffer.from([bytes[26], bytes[25], bytes[24]]), 0), pixelHeight: 1 + readUInt24BE(Buffer.from([bytes[29], bytes[28], bytes[27]]), 0) };
  }
  if ((type.includes("jpeg") || type.includes("jpg")) && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (bytes[offset] !== 0xff) { offset += 1; continue; }
      const marker = bytes[offset + 1];
      const length = bytes.readUInt16BE(offset + 2);
      if (length < 2) break;
      if ((marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)) {
        return { pixelHeight: bytes.readUInt16BE(offset + 5), pixelWidth: bytes.readUInt16BE(offset + 7) };
      }
      offset += 2 + length;
    }
  }
  return {};
}

function hasImageSignature(bytes, type) {
  if (type === "image/png") return bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  if (type === "image/jpeg") return bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;
  if (type === "image/gif") return bytes.length >= 6 && ["GIF87a", "GIF89a"].includes(bytes.subarray(0, 6).toString("ascii"));
  if (type === "image/bmp") return bytes.length >= 2 && bytes.subarray(0, 2).toString("ascii") === "BM";
  if (type === "image/webp") return bytes.length >= 12 && bytes.subarray(0, 4).toString("ascii") === "RIFF" && bytes.subarray(8, 12).toString("ascii") === "WEBP";
  return false;
}

export async function fetchImageAsset(target) {
  const targetUrl = await assertSafeOutboundUrl(target, (process.env.IMAGE_ALLOWED_TARGETS || "").split(",").map((value) => value.trim()).filter(Boolean));

  const response = await fetch(targetUrl, {
    redirect: "error",
    signal: AbortSignal.timeout(Number(process.env.OUTBOUND_TIMEOUT_MS || 15000)),
    headers: { accept: "image/png,image/jpeg,image/gif,image/webp,image/bmp;q=0.9,*/*;q=0.2", "user-agent": "Mozilla/5.0 ExpedientAIBridges", "office" },
  });
  if (!response.ok) throw new Error(`Image fetch failed: ${response.status} ${response.statusText}`);

  const type = (response.headers.get("content-type") || "application/octet-stream").split(";")[0].trim().toLowerCase();
  if (!IMAGE_TYPES.has(type)) throw new Error(`URL did not return a supported image type. Received ${type || "unknown"}.`);

  // Bound the download while streaming: arrayBuffer() would materialize an
  // arbitrarily large remote response before the size check could reject it.
  const declaredLength = Number(response.headers.get("content-length") || 0);
  if (Number.isFinite(declaredLength) && declaredLength > MAX_IMAGE_BYTES) throw new Error("Image is too large. Use an image under 8 MB.");
  const chunks = [];
  let received = 0;
  for await (const chunk of response.body || []) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    received += buffer.length;
    if (received > MAX_IMAGE_BYTES) throw new Error("Image is too large. Use an image under 8 MB.");
    chunks.push(buffer);
  }
  const bytes = Buffer.concat(chunks);
  if (!hasImageSignature(bytes, type)) throw new Error("Image content does not match its declared type.");
  const dimensions = imageDimensions(bytes, type);
  if (dimensions.pixelWidth && dimensions.pixelHeight && dimensions.pixelWidth * dimensions.pixelHeight > MAX_IMAGE_PIXELS) throw new Error("Image dimensions are too large.");

  const name = safeFilename(decodeURIComponent(targetUrl.pathname.split("/").filter(Boolean).pop() || `web-image.${extensionFromType(type)}`), `web-image.${extensionFromType(type)}`);
  return { url: targetUrl.href, name, type, size: bytes.length, ...dimensions, base64: bytes.toString("base64"), dataUrl: `data:${type};base64,${bytes.toString("base64")}` };
}

export async function handleImageAsset(req, res, url) {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "access-control-allow-methods": "GET, OPTIONS",
      "access-control-allow-headers": "content-type",
    });
    res.end();
    return;
  }
  if (req.method !== "GET") {
    res.writeHead(405, { allow: "GET, OPTIONS", "content-type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ error: { message: "Method not allowed" } }));
    return;
  }

  const target = url.searchParams.get("url")?.trim();
  if (!target) return sendJson(res, 400, { error: { message: "Missing image URL" } });
  try {
    return sendJson(res, 200, await fetchImageAsset(target));
  } catch (error) {
    return sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}