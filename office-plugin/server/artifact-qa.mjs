import zlib from "node:zlib";

const MAX_ENTRIES = 2_000;
const MAX_UNCOMPRESSED_BYTES = Number(process.env.OFFICE_ARTIFACT_MAX_UNCOMPRESSED_BYTES || 50 * 1024 * 1024);

function zipEntries(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22 || buffer.readUInt32LE(0) !== 0x04034b50) throw new Error("Generated artifact is not a valid ZIP package.");
  const entries = new Map();
  let eocd = -1;
  for (let index = buffer.length - 22; index >= Math.max(0, buffer.length - 65558); index -= 1) {
    if (buffer.readUInt32LE(index) === 0x06054b50) { eocd = index; break; }
  }
  if (eocd < 0) throw new Error("Generated artifact is missing its ZIP directory.");
  const total = buffer.readUInt16LE(eocd + 10);
  if (!total || total > MAX_ENTRIES) throw new Error("Generated artifact has an invalid entry count.");
  let offset = buffer.readUInt32LE(eocd + 16);
  let totalUncompressed = 0;
  for (let index = 0; index < total; index += 1) {
    if (offset < 0 || offset + 46 > buffer.length || buffer.readUInt32LE(offset) !== 0x02014b50) throw new Error("Generated artifact has an invalid ZIP directory entry.");
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8").replace(/\\/g, "/");
    if (!name || name.startsWith("/") || name.split("/").includes("..")) throw new Error("Generated artifact contains an unsafe package path.");
    if (localOffset + 30 > buffer.length || buffer.readUInt32LE(localOffset) !== 0x04034b50) throw new Error("Generated artifact has a dangling local ZIP entry.");
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart + compressedSize > buffer.length) throw new Error("Generated artifact has an out-of-bounds ZIP entry.");
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    totalUncompressed += uncompressedSize;
    if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) throw new Error("Generated artifact expands beyond the configured archive limit.");
    const remaining = Math.max(0, MAX_UNCOMPRESSED_BYTES - (totalUncompressed - uncompressedSize));
    let content;
    if (method === 0) content = compressed;
    else if (method === 8) content = zlib.inflateRawSync(compressed, { maxOutputLength: remaining });
    else throw new Error("Generated artifact uses an unsupported ZIP compression method.");
    if (content.length > uncompressedSize || content.length > remaining) throw new Error("Generated artifact exceeds its declared ZIP entry size.");
    entries.set(name, content);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

export function validateGeneratedOfficePackage(buffer, kind) {
  const entries = zipEntries(buffer);
  const required = ["[Content_Types].xml", "_rels/.rels"];
  if (kind === "pptx") required.push("ppt/presentation.xml", "ppt/_rels/presentation.xml.rels");
  if (kind === "docx") required.push("word/document.xml");
  if (kind === "xlsx") required.push("xl/workbook.xml");
  for (const name of required) if (!entries.has(name)) throw new Error(`Generated ${kind.toUpperCase()} is missing ${name}.`);
  return { kind, entryCount: entries.size, entries: [...entries.keys()] };
}
