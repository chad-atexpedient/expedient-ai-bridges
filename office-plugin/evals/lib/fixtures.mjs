// Build seed workbooks for eval tasks.
//
// Fixtures are generated, never committed as binaries, so the corpus stays
// diffable and license-clean. Real licensed source material must live in a
// private corpus outside this repo.

import zlib from "node:zlib";

function crc32(buffer) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = new Int32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
      table[index] = value;
    }
  }
  let crc = -1;
  for (let index = 0; index < buffer.length; index += 1) crc = (crc >>> 8) ^ table[(crc ^ buffer[index]) & 0xff];
  return (crc ^ -1) >>> 0;
}

function zip(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const [name, raw] of entries) {
    const content = Buffer.isBuffer(raw) ? raw : Buffer.from(String(raw), "utf8");
    const nameBuffer = Buffer.from(name, "utf8");
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    chunks.push(local, nameBuffer, compressed);

    const directory = Buffer.alloc(46);
    directory.writeUInt32LE(0x02014b50, 0);
    directory.writeUInt16LE(20, 4);
    directory.writeUInt16LE(20, 6);
    directory.writeUInt16LE(8, 10);
    directory.writeUInt32LE(crc, 16);
    directory.writeUInt32LE(compressed.length, 20);
    directory.writeUInt32LE(content.length, 24);
    directory.writeUInt16LE(nameBuffer.length, 28);
    directory.writeUInt32LE(offset, 42);
    central.push(directory, nameBuffer);

    offset += local.length + nameBuffer.length + compressed.length;
  }

  const centralBuffer = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuffer.length, 12);
  end.writeUInt32LE(offset, 16);

  return Buffer.concat([...chunks, centralBuffer, end]);
}

function escapeXml(value = "") {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function columnName(index) {
  let column = Math.max(0, index);
  let name = "";
  do {
    name = String.fromCharCode(65 + (column % 26)) + name;
    column = Math.floor(column / 26) - 1;
  } while (column >= 0);
  return name;
}

function cellXml(value, rowIndex, columnIndex) {
  const ref = `${columnName(columnIndex)}${rowIndex + 1}`;
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "boolean") return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
  if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
  const text = String(value);
  // A leading = marks a formula. Cached values are intentionally omitted:
  // the grader compares formulas, and a recalc oracle supplies values.
  if (text.startsWith("=")) return `<c r="${ref}"><f>${escapeXml(text.slice(1))}</f></c>`;
  return `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c>`;
}

function sheetXml(rows = []) {
  const body = rows.map((row, rowIndex) => {
    const cells = (Array.isArray(row) ? row : [row]).map((cell, columnIndex) => cellXml(cell, rowIndex, columnIndex)).join("");
    return cells ? `<row r="${rowIndex + 1}">${cells}</row>` : "";
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${body}</sheetData></worksheet>`;
}

/**
 * Build a minimal but valid .xlsx from { name, rows } sheet definitions.
 * Values may be numbers, strings, booleans, or "=FORMULA" strings.
 */
export function buildWorkbook(sheets = []) {
  const list = sheets.length ? sheets : [{ name: "Sheet1", rows: [] }];
  const entries = [];

  entries.push(["[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${list.map((_sheet, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`]);

  entries.push(["_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`]);

  entries.push(["xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${list.map((sheet, index) => `<sheet name="${escapeXml(sheet.name || `Sheet${index + 1}`)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`]);

  entries.push(["xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map((_sheet, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}</Relationships>`]);

  list.forEach((sheet, index) => entries.push([`xl/worksheets/sheet${index + 1}.xml`, sheetXml(sheet.rows || [])]));

  return zip(entries);
}

/**
 * Apply cell edits to sheet definitions, returning a new workbook.
 * Used to synthesize "what an agent did" without driving a live Excel host.
 */
export function applyEdits(sheets, edits = {}) {
  const clone = sheets.map((sheet) => ({ name: sheet.name, rows: (sheet.rows || []).map((row) => [...(Array.isArray(row) ? row : [row])]) }));

  for (const [ref, value] of Object.entries(edits)) {
    const bang = ref.lastIndexOf("!");
    const sheetName = bang >= 0 ? ref.slice(0, bang) : clone[0]?.name;
    const address = (bang >= 0 ? ref.slice(bang + 1) : ref).toUpperCase();
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) continue;

    const columnIndex = match[1].split("").reduce((sum, char) => sum * 26 + (char.charCodeAt(0) - 64), 0) - 1;
    const rowIndex = Number(match[2]) - 1;
    const sheet = clone.find((item) => item.name === sheetName);
    if (!sheet) continue;

    while (sheet.rows.length <= rowIndex) sheet.rows.push([]);
    const row = sheet.rows[rowIndex];
    while (row.length <= columnIndex) row.push("");
    row[columnIndex] = value;
  }

  return clone;
}