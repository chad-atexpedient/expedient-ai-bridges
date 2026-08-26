// Read an .xlsx package into a plain cell map.
//
// Deliberately dependency-free and read-only: the grader must not rely on the
// same code that generated a workbook, or a writer bug would grade itself as
// correct. This parses the OOXML package directly.

import zlib from "node:zlib";

const MAX_ENTRIES = 2_000;
const MAX_UNCOMPRESSED_BYTES = Number(process.env.EVAL_MAX_UNCOMPRESSED_BYTES || 50 * 1024 * 1024);

/** Minimal, bounded ZIP central-directory reader. */
export function readZipEntries(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22 || buffer.readUInt32LE(0) !== 0x04034b50) {
    throw new Error("Workbook is not a valid ZIP package.");
  }
  let eocd = -1;
  for (let index = buffer.length - 22; index >= Math.max(0, buffer.length - 65558); index -= 1) {
    if (buffer.readUInt32LE(index) === 0x06054b50) { eocd = index; break; }
  }
  if (eocd < 0) throw new Error("Workbook is missing its ZIP directory.");

  const total = buffer.readUInt16LE(eocd + 10);
  if (!total || total > MAX_ENTRIES) throw new Error("Workbook has an invalid entry count.");

  const entries = new Map();
  let offset = buffer.readUInt32LE(eocd + 16);
  let totalUncompressed = 0;

  for (let index = 0; index < total; index += 1) {
    if (offset < 0 || offset + 46 > buffer.length || buffer.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error("Workbook has an invalid ZIP directory entry.");
    }
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8").replace(/\\/g, "/");

    if (!name || name.startsWith("/") || name.split("/").includes("..")) throw new Error("Workbook contains an unsafe package path.");
    if (localOffset + 30 > buffer.length || buffer.readUInt32LE(localOffset) !== 0x04034b50) throw new Error("Workbook has a dangling local ZIP entry.");

    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart + compressedSize > buffer.length) throw new Error("Workbook has an out-of-bounds ZIP entry.");

    totalUncompressed += uncompressedSize;
    if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) throw new Error("Workbook expands beyond the configured archive limit.");
    const remaining = Math.max(0, MAX_UNCOMPRESSED_BYTES - (totalUncompressed - uncompressedSize));

    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    let content;
    if (method === 0) content = compressed;
    else if (method === 8) content = zlib.inflateRawSync(compressed, { maxOutputLength: remaining });
    else throw new Error("Workbook uses an unsupported ZIP compression method.");

    entries.set(name, content);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function decodeXmlText(value = "") {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_match, code) => String.fromCharCode(Number(code)))
    .replace(/&amp;/g, "&");
}

function parseSharedStrings(xmlText = "") {
  const strings = [];
  const items = xmlText.match(/<si\b[\s\S]*?<\/si>/g) || [];
  for (const item of items) {
    const runs = item.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
    strings.push(runs.map((run) => decodeXmlText(run.replace(/<t[^>]*>([\s\S]*?)<\/t>/, "$1"))).join(""));
  }
  return strings;
}

/** Map sheet name -> worksheet part path, preserving workbook order. */
function parseSheetIndex(workbookXml = "", relsXml = "") {
  const relationships = new Map();
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = match[0];
    const id = tag.match(/Id="([^"]+)"/)?.[1];
    const target = tag.match(/Target="([^"]+)"/)?.[1];
    if (id && target) relationships.set(id, target.replace(/^\/?xl\//, "").replace(/^\//, ""));
  }

  const sheets = [];
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const tag = match[0];
    const name = decodeXmlText(tag.match(/name="([^"]*)"/)?.[1] ?? "");
    const relId = tag.match(/r:id="([^"]+)"/)?.[1];
    const target = relId ? relationships.get(relId) : null;
    if (name) sheets.push({ name, path: target ? `xl/${target}` : null });
  }
  return sheets;
}

function parseCells(sheetXml = "", sharedStrings = []) {
  const cells = new Map();
  for (const match of sheetXml.matchAll(/<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const attributes = match[1] || "";
    const body = match[2] || "";
    const ref = attributes.match(/r="([A-Z]+\d+)"/)?.[1];
    if (!ref) continue;
    const type = attributes.match(/t="([^"]+)"/)?.[1] || "n";

    const formulaMatch = body.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    const formula = formulaMatch ? `=${decodeXmlText(formulaMatch[1])}` : "";

    let value = "";
    if (type === "inlineStr") {
      const runs = body.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
      value = runs.map((run) => decodeXmlText(run.replace(/<t[^>]*>([\s\S]*?)<\/t>/, "$1"))).join("");
    } else {
      const raw = body.match(/<v[^>]*>([\s\S]*?)<\/v>/)?.[1];
      if (raw !== undefined) {
        const decoded = decodeXmlText(raw);
        if (type === "s") value = sharedStrings[Number(decoded)] ?? "";
        else if (type === "b") value = decoded === "1";
        else if (type === "str" || type === "e") value = decoded;
        else value = decoded === "" ? "" : Number(decoded);
      }
    }

    if (value === "" && !formula) continue;
    cells.set(ref, { value, formula });
  }
  return cells;
}

/**
 * Read a workbook into { sheetName: Map<address, {value, formula}> }.
 * Empty cells are omitted entirely so diffs stay sparse.
 */
export function readWorkbook(buffer) {
  const entries = readZipEntries(buffer);
  const text = (name) => entries.has(name) ? entries.get(name).toString("utf8") : "";

  const sharedStrings = parseSharedStrings(text("xl/sharedStrings.xml"));
  const sheetIndex = parseSheetIndex(text("xl/workbook.xml"), text("xl/_rels/workbook.xml.rels"));

  const sheets = new Map();
  sheetIndex.forEach((sheet, order) => {
    // Fall back to positional naming when relationships are absent.
    const fallback = `xl/worksheets/sheet${order + 1}.xml`;
    const path = sheet.path && entries.has(sheet.path) ? sheet.path : fallback;
    if (!entries.has(path)) return;
    sheets.set(sheet.name, parseCells(text(path), sharedStrings));
  });

  return { sheets, sheetNames: [...sheets.keys()] };
}

/** All addresses present in either workbook, for union-based diffing. */
export function unionAddresses(a, b) {
  return [...new Set([...(a?.keys() ?? []), ...(b?.keys() ?? [])])].sort(compareAddresses);
}

export function compareAddresses(left, right) {
  const parse = (ref) => {
    const match = String(ref).match(/^([A-Z]+)(\d+)$/) || [];
    const letters = match[1] || "A";
    const column = letters.split("").reduce((sum, char) => sum * 26 + (char.charCodeAt(0) - 64), 0);
    return { column, row: Number(match[2] || 0) };
  };
  const a = parse(left);
  const b = parse(right);
  return a.row - b.row || a.column - b.column;
}