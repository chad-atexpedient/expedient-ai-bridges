import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import { readdir, rm } from "node:fs/promises";
import crypto from "node:crypto";
import path from "node:path";
import zlib from "node:zlib";
import { fetchImageAsset } from "./image-asset.mjs";
import { validateGeneratedOfficePackage } from "./artifact-qa.mjs";

const GENERATED_ROOT = path.join(process.env.LOCALAPPDATA || process.env.APPDATA || os.tmpdir(), "ExpedientAIBridges", "office", "generated");
const GENERATED_RETENTION_MS = Number(process.env.GENERATED_RETENTION_MS || 24 * 60 * 60 * 1000);
const MAX_PAYLOAD_BYTES = Number(process.env.GENERATED_OFFICE_MAX_PAYLOAD_BYTES || 20 * 1024 * 1024);
const SLIDE_W = 12192000;
const SLIDE_H = 6858000;
const PT_TO_EMU = 12700;
const MAX_ARCHIVE_ENTRIES = Number(process.env.OFFICE_ARCHIVE_MAX_ENTRIES || 1000);
const MAX_ARCHIVE_UNCOMPRESSED_BYTES = Number(process.env.OFFICE_ARCHIVE_MAX_UNCOMPRESSED_BYTES || 50 * 1024 * 1024);

function sendJson(res, status, body) {
  res.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

async function readJson(req, limitBytes = MAX_PAYLOAD_BYTES) {
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  const body = chunks.length ? Buffer.concat(chunks) : undefined;
  if (!body?.length) return null;
  if (body.length > limitBytes) throw new Error("Generated Office payload is too large.");
  return JSON.parse(body.toString("utf8"));
}

function xml(value = "") {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function safeFileName(value = "generated-deck.pptx") {
  return safeFileNameWithExtension(value, "pptx");
}

function safeFileNameWithExtension(value = "generated-file", extension = "pptx") {
  const fallback = `generated-file.${extension}`;
  const cleaned = String(value || fallback).replace(/[<>:"/\|?*\u0000-\u001f]/g, "-").replace(/\s+/g, " ").trim() || fallback;
  return cleaned.toLowerCase().endsWith(`.${extension}`) ? cleaned : `${cleaned}.${extension}`;
}

function officeContentType(fileName = "") {
  if (fileName.toLowerCase().endsWith(".docx")) return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (fileName.toLowerCase().endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

function asNumber(value, fallback) {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function emu(value, fallback) {
  return Math.round(asNumber(value, fallback) * PT_TO_EMU);
}

function crc32(buffer) {
  let crc = -1;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ -1) >>> 0;
}

function zip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, rawContent] of entries) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.isBuffer(rawContent) ? rawContent : Buffer.from(String(rawContent), "utf8");
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    localParts.push(local, nameBuffer, compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(content.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuffer);
    offset += local.length + nameBuffer.length + compressed.length;
  }
  const centralOffset = offset;
  const central = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(centralOffset, 16);
  return Buffer.concat([...localParts, central, end]);
}

function unzipEntries(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 22 || buffer.readUInt32LE(0) !== 0x04034b50) throw new Error("Office template is not a valid ZIP package.");
  const entries = new Map();
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65558); i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("The template PowerPoint file is not a valid Open XML zip package.");
  const total = buffer.readUInt16LE(eocd + 10);
  if (!total || total > MAX_ARCHIVE_ENTRIES) throw new Error("Office template contains too many archive entries.");
  let offset = buffer.readUInt32LE(eocd + 16);
  let totalUncompressed = 0;
  for (let i = 0; i < total; i += 1) {
    if (offset < 0 || offset + 46 > buffer.length || buffer.readUInt32LE(offset) !== 0x02014b50) throw new Error("Office template contains an invalid archive directory.");
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const nameEnd = offset + 46 + fileNameLength + extraLength + commentLength;
    if (nameEnd > buffer.length) throw new Error("Office template contains an out-of-bounds archive entry.");
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8").replace(/\\/g, "/");
    if (!name || name.startsWith("/") || name.split("/").includes("..")) throw new Error("Office template contains an unsafe archive path.");
    if (localOffset < 0 || localOffset + 30 > buffer.length || buffer.readUInt32LE(localOffset) !== 0x04034b50) throw new Error("Office template contains a dangling archive entry.");
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart < 0 || dataStart + compressedSize > buffer.length) throw new Error("Office template contains an out-of-bounds archive payload.");
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    totalUncompressed += uncompressedSize;
    if (totalUncompressed > MAX_ARCHIVE_UNCOMPRESSED_BYTES) throw new Error("Office template expands beyond the configured archive limit.");
    const remaining = Math.max(0, MAX_ARCHIVE_UNCOMPRESSED_BYTES - (totalUncompressed - uncompressedSize));
    const content = method === 0 ? compressed : method === 8 ? zlib.inflateRawSync(compressed, { maxOutputLength: remaining }) : null;
    if (!content) throw new Error("Office template uses an unsupported ZIP compression method.");
    if (content.length > uncompressedSize || content.length > remaining) throw new Error("Office template archive entry exceeds its declared size.");
    entries.set(name, content);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function zipFromMap(entries) {
  return zip([...entries.entries()]);
}

function decodeDocxTemplate(value) {
  if (typeof value !== "string" || !value.trim()) return null;
  const raw = Buffer.from(value.replace(/^data:[^,]+,/, ""), "base64");
  if (raw.length < 4 || raw.length > MAX_PAYLOAD_BYTES || raw.readUInt32LE(0) !== 0x04034b50) {
    throw new Error("Word template must be a valid, bounded DOCX ZIP package.");
  }
  const entries = unzipEntries(raw);
  if (!entries.has("[Content_Types].xml") || !entries.has("word/document.xml")) {
    throw new Error("Word template is missing required Open XML document parts.");
  }
  if (entries.size > 2_000 || [...entries.values()].reduce((sum, item) => sum + item.length, 0) > 50 * 1024 * 1024) {
    throw new Error("Word template contains too many or too-large package entries.");
  }
  return entries;
}

function mergeDocxContentTypes(templateXml = "", generatedXml = "") {
  const template = templateXml || generatedXml;
  const overrides = [...String(generatedXml).matchAll(/<Override\b[^>]*PartName="\/([^\"]+)"[^>]*ContentType="([^\"]+)"[^>]*\/>/g)]
    .map((match) => `<Override PartName="/${match[1]}" ContentType="${match[2]}"/>`);
  let merged = template || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`;
  for (const override of overrides) {
    const part = override.match(/PartName="\/([^\"]+)"/)?.[1];
    if (!part) continue;
    const pattern = new RegExp(`<Override\\s+[^>]*PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`, "g");
    merged = merged.replace(pattern, "");
  }
  return merged.replace("</Types>", `${overrides.join("")}</Types>`);
}

function mergeDocxRelationships(templateXml = "", generatedXml = "") {
  const template = relationshipItems(templateXml);
  const generated = relationshipItems(generatedXml);
  const byId = new Map(template.map((item) => [item.id, item]));
  for (const item of generated) byId.set(item.id, item);
  return rels([...byId.values()]);
}

function templateSectionReferences(entries) {
  const documentXml = entries?.get("word/document.xml")?.toString("utf8") || "";
  const relsXml = entries?.get("word/_rels/document.xml.rels")?.toString("utf8") || "";
  const ids = [...documentXml.matchAll(/<w:(headerReference|footerReference)\b[^>]*\br:id="([^"]+)"[^>]*\/>/g)]
    .map((match) => ({ kind: match[1], id: match[2] }))
    .filter((item) => relationshipItems(relsXml).some((rel) => rel.id === item.id && /\/header$|\/footer$/.test(rel.type)));
  return ids.map((item) => `<w:${item.kind} w:type="default" r:id="${xml(item.id)}"/>`).join("");
}

function removeMatchingEntries(entries, pattern) {
  for (const key of [...entries.keys()]) if (pattern.test(key)) entries.delete(key);
}

function templateLayoutTarget(entries) {
  const relsXml = entries.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") || "";
  const masterTarget = relsXml.match(/Type="[^"]+\/slideMaster"[^>]*Target="([^"]+)"/)?.[1];
  const masterPath = masterTarget ? `ppt/${masterTarget.replace(/^\/+/, "")}`.replace(/\/[^/]+\/\.\.\//g, "/") : "ppt/slideMasters/slideMaster1.xml";
  const masterRelsPath = masterPath.replace("ppt/slideMasters/", "ppt/slideMasters/_rels/") + ".rels";
  const masterRels = entries.get(masterRelsPath)?.toString("utf8") || "";
  const layoutTarget = masterRels.match(/Type="[^"]+\/slideLayout"[^>]*Target="([^"]+)"/)?.[1];
  if (layoutTarget) return layoutTarget.startsWith("../") ? layoutTarget : `../slideLayouts/${path.basename(layoutTarget)}`;
  return "../slideLayouts/slideLayout1.xml";
}

function relationshipItems(relsXml = "") {
  return [...String(relsXml || "").matchAll(/<Relationship\b[^>]*>/g)].map((match) => ({
    id: match[0].match(/\bId="([^"]+)"/)?.[1] || "",
    type: match[0].match(/\bType="([^"]+)"/)?.[1] || "",
    target: match[0].match(/\bTarget="([^"]+)"/)?.[1] || "",
  })).filter((rel) => rel.id || rel.type || rel.target);
}

function normalizePptPath(basePath = "ppt", target = "") {
  const raw = String(target || "").replace(/^\/+/, "");
  if (raw.startsWith("ppt/")) return raw;
  const parts = `${basePath}/${raw}`.split("/");
  const normalized = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") normalized.pop();
    else normalized.push(part);
  }
  return normalized.join("/");
}

function templateLayoutPath(layoutTarget = "../slideLayouts/slideLayout1.xml") {
  return normalizePptPath("ppt/slides", layoutTarget || "../slideLayouts/slideLayout1.xml");
}

function layoutName(layoutXml = "", fallback = "") {
  return layoutXml.match(/<p:cSld\b[^>]*\bname="([^"]+)"/)?.[1]
    || layoutXml.match(/<p:sldLayout\b[^>]*\btype="([^"]+)"/)?.[1]
    || fallback;
}

function layoutType(layoutXml = "") {
  return layoutXml.match(/<p:sldLayout\b[^>]*\btype="([^"]+)"/)?.[1] || "";
}

function templateLayouts(entries) {
  const presentationRels = entries.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") || "";
  const masterRels = relationshipItems(presentationRels).filter((rel) => /\/slideMaster$/.test(rel.type));
  const layouts = [];
  for (const masterRel of masterRels.length ? masterRels : [{ target: "slideMasters/slideMaster1.xml" }]) {
    const masterPath = normalizePptPath("ppt", masterRel.target);
    const masterRelsPath = masterPath.replace("ppt/slideMasters/", "ppt/slideMasters/_rels/") + ".rels";
    const masterRelsXml = entries.get(masterRelsPath)?.toString("utf8") || "";
    const layoutRels = relationshipItems(masterRelsXml).filter((rel) => /\/slideLayout$/.test(rel.type));
    for (const rel of layoutRels) {
      const pathName = normalizePptPath(path.posix.dirname(masterPath), rel.target);
      const xmlText = entries.get(pathName)?.toString("utf8") || "";
      if (!xmlText) continue;
      layouts.push({
        index: layouts.length,
        relId: rel.id,
        target: pathName.replace(/^ppt\/slides\//, "").replace(/^ppt\//, "../"),
        path: pathName,
        name: layoutName(xmlText, path.basename(pathName, ".xml")),
        type: layoutType(xmlText),
        xml: xmlText,
      });
    }
  }
  if (!layouts.length) {
    const fallbackPath = templateLayoutPath(templateLayoutTarget(entries));
    const xmlText = entries.get(fallbackPath)?.toString("utf8") || "";
    layouts.push({ index: 0, relId: "rLayout", target: "../slideLayouts/slideLayout1.xml", path: fallbackPath, name: layoutName(xmlText, "slideLayout1"), type: layoutType(xmlText), xml: xmlText });
  }
  return layouts;
}

function normalizedLabel(value = "") {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function inferLayoutNeed(slide = {}) {
  if (Array.isArray(slide.charts) && slide.charts.length) return "chart";
  if (Array.isArray(slide.tables) && slide.tables.length) return "table";
  if (Array.isArray(slide.images) && slide.images.length) return "picture";
  if (slide.subtitle) return "title";
  return "content";
}

function selectTemplateLayout(layouts = [], slide = {}) {
  if (!layouts.length) return null;
  if (typeof slide.layoutIndex === "number" && layouts[slide.layoutIndex]) return layouts[slide.layoutIndex];
  const wantedName = normalizedLabel(slide.layoutName || slide.layout || "");
  if (wantedName) {
    const byName = layouts.find((layout) => normalizedLabel(layout.name) === wantedName || normalizedLabel(layout.path).includes(wantedName));
    if (byName) return byName;
  }
  const wantedType = normalizedLabel(slide.layoutType || inferLayoutNeed(slide));
  const byType = layouts.find((layout) => normalizedLabel(layout.type) === wantedType || normalizedLabel(layout.name).includes(wantedType));
  if (byType) return byType;
  if (wantedType === "picture") {
    const pictureLayout = layouts.find((layout) => /picture|image|media|visual/.test(normalizedLabel(layout.name)));
    if (pictureLayout) return pictureLayout;
  }
  return layouts[0];
}

function stripContentTypeOverrides(contentTypesXml = "", prefixes = []) {
  let next = contentTypesXml;
  for (const prefix of prefixes) {
    const pattern = new RegExp(`<Override\\s+PartName="/${prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}[^>]*>`, "g");
    next = next.replace(pattern, "");
  }
  return next;
}

function addContentTypeOverrides(contentTypesXml = "", overrides = []) {
  const clean = stripContentTypeOverrides(contentTypesXml, ["ppt/slides/slide", "ppt/notesSlides/notesSlide", "ppt/charts/chart"]);
  return clean.replace("</Types>", `${overrides.join("")}</Types>`);
}

function presentationRelsWithSlides(entries, slideCount) {
  const current = entries.get("ppt/_rels/presentation.xml.rels")?.toString("utf8") || rels([]);
  const preserved = [...current.matchAll(/<Relationship\b[^>]*>/g)]
    .map((match) => match[0])
    .filter((rel) => !/\/relationships\/slide"/.test(rel));
  const slideRels = Array.from({ length: slideCount }, (_unused, index) => `<Relationship Id="rSlide${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${preserved.join("")}${slideRels.join("")}</Relationships>`;
}

function presentationWithSlides(entries, slideCount) {
  const current = entries.get("ppt/presentation.xml")?.toString("utf8");
  const sldIdLst = Array.from({ length: slideCount }, (_unused, index) => `<p:sldId id="${256 + index}" r:id="rSlide${index + 1}"/>`).join("");
  if (current?.includes("<p:sldIdLst")) return current.replace(/<p:sldIdLst[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIdLst}</p:sldIdLst>`);
  if (current?.includes("<p:sldMasterIdLst")) return current.replace(/(<\/p:sldMasterIdLst>)/, `$1<p:sldIdLst>${sldIdLst}</p:sldIdLst>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIdLst}</p:sldIdLst><p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
}

function shapeId(shapeXml = "") {
  return Number(shapeXml.match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)?.[1] || 0);
}

function placeholderKind(shapeXml = "") {
  const phTag = shapeXml.match(/<p:ph\b[^>]*>/)?.[0] || "";
  const type = phTag.match(/\btype="([^"]+)"/)?.[1] || "";
  const name = shapeXml.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1] || "";
  const haystack = `${type} ${name}`.toLowerCase();
  if (/ctrtitle|title/.test(haystack)) return "title";
  if (/subtitle/.test(haystack)) return "subtitle";
  if (/pic|image|picture|media/.test(haystack)) return "image";
  if (/chart|graph/.test(haystack)) return "chart";
  if (/tbl|table/.test(haystack)) return "table";
  if (/body|content|object|obj/.test(haystack)) return "body";
  return "other";
}

function shapeBounds(shapeXml = "") {
  const match = shapeXml.match(/<a:off\b[^>]*\bx="(-?\d+)"[^>]*\by="(-?\d+)"[\s\S]*?<a:ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
  if (!match) return null;
  return { x: Number(match[1]), y: Number(match[2]), cx: Number(match[3]), cy: Number(match[4]) };
}

function placementFromBounds(bounds) {
  if (!bounds) return null;
  return {
    left: bounds.x / PT_TO_EMU,
    top: bounds.y / PT_TO_EMU,
    width: bounds.cx / PT_TO_EMU,
    height: bounds.cy / PT_TO_EMU,
  };
}

function replaceShapeText(shapeXml = "", text = "") {
  const txBody = `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${textRuns(text)}</p:txBody>`;
  if (/<p:txBody>[\s\S]*?<\/p:txBody>/.test(shapeXml)) return shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, txBody);
  return shapeXml.replace(/<\/p:sp>\s*$/, `${txBody}</p:sp>`);
}

function placeholderShapesFromLayout(layoutXml = "") {
  return [...String(layoutXml || "").matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0])
    .filter((shape) => shape.includes("<p:ph"))
    .map((shape) => ({ kind: placeholderKind(shape), id: shapeId(shape), bounds: shapeBounds(shape), xml: shape }));
}

function placeholderPlacements(placeholders = []) {
  const buckets = { image: [], chart: [], table: [], content: [] };
  for (const placeholder of placeholders) {
    const placement = placementFromBounds(placeholder.bounds);
    if (!placement) continue;
    if (placeholder.kind === "image") buckets.image.push(placement);
    else if (placeholder.kind === "chart") buckets.chart.push(placement);
    else if (placeholder.kind === "table") buckets.table.push(placement);
    else if (placeholder.kind === "body" || placeholder.kind === "other") buckets.content.push(placement);
  }
  return buckets;
}

function applyPlacement(item = {}, placement) {
  if (!placement || !item || typeof item !== "object") return item;
  return {
    ...item,
    left: item.left ?? placement.left,
    top: item.top ?? placement.top,
    width: item.width ?? placement.width,
    height: item.height ?? placement.height,
  };
}

function applyTemplatePlacements(slide = {}, placements = {}) {
  const imageTargets = [...(placements.image || []), ...(placements.content || [])];
  const chartTargets = [...(placements.chart || []), ...(placements.content || [])];
  const tableTargets = [...(placements.table || []), ...(placements.content || [])];
  const images = Array.isArray(slide.images) ? slide.images.map((image, index) => applyPlacement(image, imageTargets[index])) : slide.images;
  const charts = Array.isArray(slide.charts) ? slide.charts.map((chart, index) => applyPlacement(chart, chartTargets[index])) : slide.charts;
  const tables = Array.isArray(slide.tables) ? slide.tables.map((table, index) => applyPlacement(table, tableTargets[index])) : slide.tables;
  return { ...slide, images, charts, tables };
}

const DEFAULT_BRAND = {
  dark: "111827", light: "FFFFFF", accent1: "2563EB", accent2: "7C3AED", accent3: "0F766E",
  majorFont: "Aptos Display", minorFont: "Aptos",
};

function normalizedBrandProfile(payload = {}) {
  const profile = payload?.brandProfile || payload?.template?.brandProfile;
  const colors = profile?.colors && typeof profile.colors === "object" ? profile.colors : {};
  const fonts = profile?.fonts && typeof profile.fonts === "object" ? profile.fonts : {};
  const color = (key, fallback) => hexColor(colors[key], fallback);
  const font = (value, fallback) => typeof value === "string" && value.trim() && value.length <= 120 ? value.trim() : fallback;
  return {
    dark: color("dk1", DEFAULT_BRAND.dark), light: color("lt1", DEFAULT_BRAND.light),
    accent1: color("accent1", DEFAULT_BRAND.accent1), accent2: color("accent2", DEFAULT_BRAND.accent2),
    accent3: color("accent3", DEFAULT_BRAND.accent3), majorFont: font(fonts.major, DEFAULT_BRAND.majorFont),
    minorFont: font(fonts.minor, DEFAULT_BRAND.minorFont),
  };
}

function generatedSlideObjects(startId, slide, imageRels, chartRels = [], hyperlinkRels = []) {
  let nextId = startId;
  const generatedShapes = Array.isArray(slide.shapes) ? slide.shapes.slice(0, 24) : [];
  const generatedTables = Array.isArray(slide.tables) ? slide.tables.slice(0, 8) : [];
  const objects = [];
  for (const [index, shape] of generatedShapes.entries()) objects.push(generatedShape(nextId++, shape, index, slide.__brand));
  for (const [index, table] of generatedTables.entries()) objects.push(generatedTable(nextId++, table, index, slide.__brand));
  for (const [index, { relId, chart }] of chartRels.entries()) objects.push(generatedChartFrame(nextId++, relId, chart, index));
  for (const [index, { relId, image }] of imageRels.entries()) objects.push(picture(nextId++, relId, image, index));
  for (const [index, { relId, link }] of hyperlinkRels.entries()) objects.push(pptHyperlinkBox(nextId++, link, relId, index));
  objects.push(generatedSlideChrome(nextId, slide));
  return objects.join("");
}

function deckChromeSlide(rawSlide = {}, payload = {}, index = 0, total = 1) {
  const slide = rawSlide && typeof rawSlide === "object" ? rawSlide : {};
  return {
    ...slide,
    footer: slide.footer ?? slide.footerText ?? payload.footer ?? payload.footerText,
    dateText: slide.dateText ?? slide.date ?? payload.dateText ?? payload.date,
    showDate: slide.showDate ?? payload.showDate,
    showFooter: slide.showFooter ?? payload.showFooter,
    showSlideNumber: slide.showSlideNumber ?? slide.showSlideNumbers ?? payload.showSlideNumber ?? payload.showSlideNumbers,
    slideNumberText: slide.slideNumberText ?? payload.slideNumberText,
    slideNumberFormat: slide.slideNumberFormat ?? payload.slideNumberFormat,
    confidentialityLabel: slide.confidentialityLabel ?? payload.confidentialityLabel,
    footerColor: slide.footerColor ?? payload.footerColor,
    footerFontSize: slide.footerFontSize ?? payload.footerFontSize,
    __slideIndex: index,
    __slideCount: total,
    __brand: normalizedBrandProfile(payload),
  };
}

function contentTypeWithDefault(contentTypesXml = "", ext = "png", type = imageContentType(ext)) {
  if (new RegExp(`<Default\\s+Extension="${ext}"`, "i").test(contentTypesXml)) return contentTypesXml;
  return contentTypesXml.replace("</Types>", `<Default Extension="${xml(ext)}" ContentType="${xml(type)}"/></Types>`);
}

function createTemplatePreservingPptx(payload = {}) {
  const templateBase64 = payload.template?.base64 || payload.templateBase64;
  if (!templateBase64 || typeof templateBase64 !== "string") return null;
  const rawSlides = Array.isArray(payload.slides) && payload.slides.length ? payload.slides.slice(0, 100) : [{ title: payload.title || "Generated deck", body: "" }];
  const slides = rawSlides.map((slide, index) => deckChromeSlide(slide, payload, index, rawSlides.length));
  const brand = normalizedBrandProfile(payload);
  const entries = unzipEntries(Buffer.from(templateBase64.replace(/^data:[^,]+,/, ""), "base64"));
  if (!entries.has("ppt/presentation.xml")) throw new Error("The template PowerPoint file is missing ppt/presentation.xml.");

  removeMatchingEntries(entries, /^ppt\/slides\/slide\d+\.xml$/);
  removeMatchingEntries(entries, /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/);
  removeMatchingEntries(entries, /^ppt\/notesSlides\/notesSlide\d+\.xml$/);
  removeMatchingEntries(entries, /^ppt\/charts\/chart\d+\.xml$/);

  const layouts = templateLayouts(entries);
  const imageExts = new Set();
  let mediaIndex = Math.max(1, ...[...entries.keys()].map((key) => Number(key.match(/^ppt\/media\/image(\d+)\./)?.[1] || 0))) + 1;
  let chartIndex = 1;
  let embeddedWorkbookIndex = 1;
  let noteCount = 0;

  for (const [index, rawSlide] of slides.entries()) {
    const layout = selectTemplateLayout(layouts, rawSlide) || layouts[0];
    const layoutTarget = layout?.target || templateLayoutTarget(entries);
    const layoutXml = layout?.xml || entries.get(templateLayoutPath(layoutTarget))?.toString("utf8") || "";
    const layoutPlaceholders = placeholderShapesFromLayout(layoutXml);
    const objectPlacements = placeholderPlacements(layoutPlaceholders);
    const slide = applyTemplatePlacements(rawSlide, objectPlacements);
    const slideNumber = index + 1;
    const slideRels = [{ id: "rLayout", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: layoutTarget }];
    const imageRels = [];
    const chartRels = [];
    const hyperlinkRels = [];
    const images = Array.isArray(slide.images) ? slide.images.slice(0, 8) : [];
    for (const image of images) {
      if (!image?.base64) continue;
      const ext = imageExt(image.type);
      imageExts.add(ext);
      const mediaName = `image${mediaIndex++}.${ext}`;
      const relId = `rImg${imageRels.length + 1}`;
      entries.set(`ppt/media/${mediaName}`, embeddedImageBytes(image));
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `../media/${mediaName}` });
      imageRels.push({ relId, image: { ...image, name: image.name || mediaName } });
    }
    const charts = Array.isArray(slide.charts) ? slide.charts.slice(0, 6) : [];
    for (const chart of charts) {
      if (!chart || typeof chart !== "object") continue;
      const styledChart = brandChart(chart, slide.__brand);
      const chartNumber = chartIndex++;
      const workbookName = `chartData${embeddedWorkbookIndex++}.xlsx`;
      const relId = `rChart${chartNumber}`;
      entries.set(`ppt/charts/chart${chartNumber}.xml`, chartXml(styledChart, { externalWorkbook: true }));
      entries.set(`ppt/charts/_rels/chart${chartNumber}.xml.rels`, chartRelsXml(workbookName));
      entries.set(`ppt/embeddings/${workbookName}`, embeddedChartWorkbook(styledChart));
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: `../charts/chart${chartNumber}.xml` });
      chartRels.push({ relId, chart: styledChart });
    }
    const links = Array.isArray(slide.links) ? slide.links.slice(0, 12) : [];
    for (const link of links) {
      if (!link || typeof link !== "object" || !link.url) continue;
      const relId = `rLink${hyperlinkRels.length + 1}`;
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: String(link.url), targetMode: "External" });
      hyperlinkRels.push({ relId, link });
    }
    if (slide.notes) {
      noteCount += 1;
      entries.set(`ppt/notesSlides/notesSlide${slideNumber}.xml`, notesXml(slide.notes));
      slideRels.push({ id: "rNotes", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", target: `../notesSlides/notesSlide${slideNumber}.xml` });
    }
    entries.set(`ppt/slides/slide${slideNumber}.xml`, templateSlideXml(slide, index, layoutXml, imageRels, chartRels, hyperlinkRels));
    entries.set(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, rels(slideRels));
  }

  entries.set("ppt/presentation.xml", presentationWithSlides(entries, slides.length));
  entries.set("ppt/_rels/presentation.xml.rels", presentationRelsWithSlides(entries, slides.length));
  const overrides = [
    ...slides.map((_slide, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`),
    ...Array.from({ length: noteCount }, (_unused, index) => `<Override PartName="/ppt/notesSlides/notesSlide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`),
    ...Array.from({ length: chartIndex - 1 }, (_unused, index) => `<Override PartName="/ppt/charts/chart${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`),
    ...Array.from({ length: embeddedWorkbookIndex - 1 }, (_unused, index) => `<Override PartName="/ppt/embeddings/chartData${index + 1}.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`),
  ];
  let types = entries.get("[Content_Types].xml")?.toString("utf8") || contentTypes(slides.length, noteCount, imageExts, chartIndex - 1);
  types = addContentTypeOverrides(types, overrides);
  for (const ext of imageExts) types = contentTypeWithDefault(types, ext);
  entries.set("[Content_Types].xml", types);
  return zipFromMap(entries);
}

function textRuns(text) {
  return String(text || "").split(/\r?\n/).map((line) => `<a:p><a:r><a:rPr lang="en-US"/><a:t>${xml(line)}</a:t></a:r></a:p>`).join("") || "<a:p/>";
}

function textBox(id, name, text, x, y, cx, cy, fontSize = 1800, bold = false) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${fontSize}"${bold ? ' b="1"' : ""}/><a:t>${xml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}


function pptHyperlinkBox(id, link = {}, relId, index = 0) {
  const x = emu(link.left, 52);
  const y = emu(link.top, 470 + index * 34);
  const cx = emu(link.width, 560);
  const cy = emu(link.height, 30);
  const fontSize = Math.round(asNumber(link.fontSize, 1300));
  const text = link.text || link.label || link.url || `Link ${index + 1}`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xml(text)}"><a:hlinkClick r:id="${relId}" tooltip="${xml(link.url || "")}"/></p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${fontSize}" u="sng"><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:rPr><a:t>${xml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function chromeTextBox(id, name, text, xPt, yPt, wPt, hPt, options = {}) {
  if (!text) return "";
  const fontSize = Math.round(asNumber(options.fontSize, 900));
  const color = hexColor(options.color || "64748B", "64748B");
  const align = pptAlign(options.align || "left");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(xPt, 0)}" y="${emu(yPt, 0)}"/><a:ext cx="${emu(wPt, 100)}" cy="${emu(hPt, 16)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="none"/><a:lstStyle/><a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="en-US" sz="${fontSize}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${xml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function generatedSlideChrome(startId, slide = {}) {
  let nextId = startId;
  const footerText = slide.showFooter === false ? "" : String(slide.footer || "").trim();
  const dateText = slide.showDate === false ? "" : String(slide.dateText || "").trim();
  const label = String(slide.confidentialityLabel || "").trim();
  const showSlideNumber = slide.showSlideNumber === true || slide.showSlideNumbers === true;
  const number = (slide.__slideIndex || 0) + 1;
  const count = slide.__slideCount || number;
  const numberFormat = String(slide.slideNumberFormat || "{n}").trim() || "{n}";
  const slideNumberText = showSlideNumber ? String(slide.slideNumberText || numberFormat).replace(/\{n\}/g, String(number)).replace(/\{total\}/g, String(count)) : "";
  const color = slide.footerColor || slide.__brand?.dark || "64748B";
  const fontSize = slide.footerFontSize || 850;
  return [
    chromeTextBox(nextId++, "Footer", footerText, 42, 510, 315, 18, { color, fontSize, align: "left" }),
    chromeTextBox(nextId++, "Date", dateText, 365, 510, 130, 18, { color, fontSize, align: "center" }),
    chromeTextBox(nextId++, "Confidentiality", label, 500, 510, 120, 18, { color, fontSize, align: "center" }),
    chromeTextBox(nextId++, "Slide Number", slideNumberText, 622, 510, 60, 18, { color, fontSize, align: "right" }),
  ].join("");
}

function bodyBox(id, name, text, x, y, cx, cy) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${xml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${textRuns(text)}</p:txBody></p:sp>`;
}

function hexColor(value = "", fallback = "FFFFFF") {
  const raw = String(value || "").trim().replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(raw) ? raw : fallback;
}

function solidFill(color = "") {
  return color ? `<a:solidFill><a:srgbClr val="${hexColor(color)}"/></a:solidFill>` : "<a:noFill/>";
}

function slideBackground(slide = {}) {
  const color = slide.backgroundColor || slide.background?.color;
  if (!color) return "";
  return `<p:bg><p:bgPr>${solidFill(color)}<a:effectLst/></p:bgPr></p:bg>`;
}

function pptPresetShape(value = "rectangle") {
  const map = { rectangle: "rect", roundedRectangle: "roundRect", oval: "ellipse", triangle: "triangle", diamond: "diamond", pentagon: "pentagon", hexagon: "hexagon", cloud: "cloud", line: "rect" };
  return map[value] || map.rectangle;
}

function generatedShape(id, shape = {}, index = 0, brand = DEFAULT_BRAND) {
  const x = emu(shape.left, 70 + index * 24);
  const y = emu(shape.top, 300 + index * 28);
  const cx = emu(shape.width, 180);
  const cy = emu(shape.height, 72);
  const name = xml(shape.name || shape.text || `Shape ${index + 1}`);
  const fill = solidFill(shape.fillColor || brand?.accent1 || "E0F2FE");
  const line = `<a:ln><a:solidFill><a:srgbClr val="${hexColor(shape.lineColor, brand?.accent2 || "2563EB")}"/></a:solidFill></a:ln>`;
  const text = shape.text ? `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${Math.round(asNumber(shape.fontSize, 1400))}"><a:latin typeface="${xml(shape.fontFamily || brand?.minorFont || "Aptos")}"/></a:rPr><a:t>${xml(shape.text)}</a:t></a:r></a:p></p:txBody>` : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${pptPresetShape(shape.shapeType)}"><a:avLst/></a:prstGeom>${fill}${line}</p:spPr>${text}</p:sp>`;
}

function pptAlign(value = "") {
  const map = { left: "l", center: "ctr", centre: "ctr", right: "r", justify: "just", l: "l", ctr: "ctr", r: "r", just: "just" };
  return map[String(value || "").toLowerCase()] || "l";
}

function pptVerticalAnchor(value = "") {
  const map = { top: "t", middle: "ctr", center: "ctr", centre: "ctr", bottom: "b", t: "t", ctr: "ctr", b: "b" };
  return map[String(value || "").toLowerCase()] || "ctr";
}

function tableBorderXml(color = "CBD5E1", width = 0.75) {
  if (color === null || color === false || String(color || "").toLowerCase() === "none") return "";
  const lineWidth = Math.max(0, Math.round(asNumber(width, 0.75) * PT_TO_EMU));
  const line = `<a:ln w="${lineWidth}"><a:solidFill><a:srgbClr val="${hexColor(color, "CBD5E1")}"/></a:solidFill></a:ln>`;
  return `<a:lnL>${line}</a:lnL><a:lnR>${line}</a:lnR><a:lnT>${line}</a:lnT><a:lnB>${line}</a:lnB>`;
}

function tableCellValue(cell) {
  if (cell && typeof cell === "object" && !Array.isArray(cell)) return cell.text ?? cell.value ?? "";
  return cell ?? "";
}

function tableCell(cell = "", options = {}) {
  const cellOptions = cell && typeof cell === "object" && !Array.isArray(cell) ? cell : {};
  const fillColor = cellOptions.fillColor ?? options.fillColor;
  const textColor = cellOptions.textColor ?? options.textColor;
  const fontSize = Math.round(asNumber(cellOptions.fontSize, options.fontSize || 1200));
  const bold = cellOptions.bold ?? options.bold;
  const align = pptAlign(cellOptions.align ?? options.align);
  const verticalAlign = pptVerticalAnchor(cellOptions.verticalAlign ?? options.verticalAlign);
  const fill = fillColor ? solidFill(fillColor) : "";
  const textFill = textColor ? `<a:solidFill><a:srgbClr val="${hexColor(textColor, "111827")}"/></a:solidFill>` : "";
  const border = tableBorderXml(options.borderColor, options.borderWidth);
  return `<a:tc><a:txBody><a:bodyPr anchor="${verticalAlign}" lIns="45720" rIns="45720" tIns="22860" bIns="22860"/><a:lstStyle/><a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="en-US" sz="${fontSize}"${bold ? ' b="1"' : ""}>${textFill}</a:rPr><a:t>${xml(tableCellValue(cell))}</a:t></a:r></a:p></a:txBody><a:tcPr>${fill}${border}</a:tcPr></a:tc>`;
}

function generatedTable(id, table = {}, index = 0, brand = DEFAULT_BRAND) {
  const rows = Array.isArray(table.values) ? table.values.map((row) => Array.isArray(row) ? row : [row]) : [];
  if (!rows.length) return "";
  const columnCount = Math.max(1, ...rows.map((row) => row.length));
  const x = emu(table.left, 60);
  const y = emu(table.top, 150 + index * 40);
  const cx = emu(table.width, 600);
  const cy = emu(table.height, Math.max(120, rows.length * 28));
  const columnWidths = Array.isArray(table.columnWidths) && table.columnWidths.length ? table.columnWidths : [];
  const rowHeights = Array.isArray(table.rowHeights) && table.rowHeights.length ? table.rowHeights : [];
  const columnWidth = Math.floor(cx / columnCount);
  const rowHeight = Math.floor(cy / rows.length);
  const headerFill = table.headerFillColor || brand?.accent1 || "DBEAFE";
  const grid = Array.from({ length: columnCount }, (_unused, colIndex) => `<a:gridCol w="${columnWidths[colIndex] ? emu(columnWidths[colIndex], table.width ? table.width / columnCount : 600 / columnCount) : columnWidth}"/>`).join("");
  const tr = rows.map((row, rowIndex) => {
    const isHeader = rowIndex === 0 && table.firstRow !== false;
    const isBand = !isHeader && table.bandRows !== false && rowIndex % 2 === 0;
    const fillColor = isHeader ? headerFill : isBand ? table.bandFillColor || table.alternateRowFillColor || "F8FAFC" : table.bodyFillColor || "";
    const cellOptions = {
      fillColor,
      textColor: isHeader ? table.headerTextColor || brand?.light || "FFFFFF" : table.textColor || brand?.dark || "111827",
      fontSize: isHeader ? table.headerFontSize || table.fontSize || 1200 : table.fontSize || 1150,
      bold: isHeader ? table.headerBold !== false : table.bold === true,
      align: isHeader ? table.headerAlign || table.align || "center" : table.align || "left",
      verticalAlign: table.verticalAlign || "middle",
      borderColor: table.borderColor || brand?.accent2 || "CBD5E1",
      borderWidth: table.borderWidth || 0.75,
    };
    const height = rowHeights[rowIndex] ? emu(rowHeights[rowIndex], table.height ? table.height / rows.length : 28) : rowHeight;
    return `<a:tr h="${height}">${Array.from({ length: columnCount }, (_unused, colIndex) => tableCell(row[colIndex] ?? "", cellOptions)).join("")}</a:tr>`;
  }).join("");
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${xml(table.name || `Table ${index + 1}`)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${table.firstRow === false ? 0 : 1}" bandRow="${table.bandRows === false ? 0 : 1}"/><a:tblGrid>${grid}</a:tblGrid>${tr}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function chartTypeTag(type = "bar") {
  const requested = String(type || "bar");
  if (requested === "line") return "c:lineChart";
  if (requested === "pie") return "c:pieChart";
  if (requested === "area") return "c:areaChart";
  if (requested === "doughnut") return "c:doughnutChart";
  if (requested === "scatter") return "c:scatterChart";
  if (requested === "combo") return "combo";
  return "c:barChart";
}

function chartLegendPosition(value = "r") {
  const map = { right: "r", left: "l", top: "t", bottom: "b", r: "r", l: "l", t: "t", b: "b" };
  return map[String(value || "r").toLowerCase()] || "r";
}

function chartRichText(text = "") {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${xml(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

function chartAxisTitle(text = "") {
  return text ? `<c:title>${chartRichText(text)}<c:layout/></c:title>` : "";
}

function chartDataLabels(chart = {}) {
  if (!chart.dataLabels) return "";
  const percent = chart.chartType === "pie" || chart.chartType === "doughnut" ? 1 : 0;
  return `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${percent}"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

function chartSeriesShape(chart = {}, index = 0) {
  const colors = Array.isArray(chart.colors) && chart.colors.length ? chart.colors : [];
  const color = colors[index % colors.length];
  return color ? `<c:spPr><a:solidFill><a:srgbClr val="${hexColor(color, "2563EB")}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${hexColor(color, "2563EB")}"/></a:solidFill></a:ln></c:spPr>` : "";
}

function normalizeChartSeries(chart = {}) {
  if (Array.isArray(chart.series) && chart.series.length) {
    const categories = Array.isArray(chart.categories) && chart.categories.length
      ? chart.categories.map((category, index) => String(category ?? `Item ${index + 1}`)).slice(0, 50)
      : Array.isArray(chart.series[0]?.values)
        ? chart.series[0].values.map((_value, index) => `Item ${index + 1}`).slice(0, 50)
        : [];
    const series = chart.series.slice(0, 12).map((item, index) => ({
      name: item?.name || item?.seriesName || `Series ${index + 1}`,
      sourceIndex: index,
      values: (Array.isArray(item?.values) ? item.values : []).map((value) => Number(value) || 0).slice(0, categories.length || 50),
    }));
    const width = Math.max(categories.length, ...series.map((item) => item.values.length), 0);
    return {
      categories: categories.length ? categories : Array.from({ length: width }, (_unused, index) => `Item ${index + 1}`),
      series: series.map((item) => ({ ...item, values: Array.from({ length: width }, (_unused, index) => Number(item.values[index]) || 0) })),
    };
  }

  const rows = Array.isArray(chart.values) ? chart.values : [];
  const categories = rows.map((row, index) => Array.isArray(row) ? String(row[0] ?? `Item ${index + 1}`) : `Item ${index + 1}`).slice(0, 50);
  const values = rows.map((row) => Array.isArray(row) ? Number(row[1] ?? 0) || 0 : Number(row) || 0).slice(0, 50);
  return { categories, series: [{ name: chart.seriesName || chart.title || "Series 1", sourceIndex: 0, values }] };
}

function brandChart(chart = {}, brand = DEFAULT_BRAND) {
  return { ...chart, colors: Array.isArray(chart.colors) && chart.colors.length ? chart.colors : [brand?.accent1, brand?.accent2, brand?.accent3].filter(Boolean) };
}

function chartSheetRefName(name = "Chart Data") {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function chartCellFormula(columnIndex, rowIndex, sheetName = "Chart Data") {
  return `${chartSheetRefName(sheetName)}!$${columnName(columnIndex)}$${rowIndex}`;
}

function chartRangeFormula(startColumnIndex, startRowIndex, endColumnIndex, endRowIndex, sheetName = "Chart Data") {
  return `${chartSheetRefName(sheetName)}!$${columnName(startColumnIndex)}$${startRowIndex}:$${columnName(endColumnIndex)}$${endRowIndex}`;
}

function chartStringReference(formula = "", values = []) {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${xml(value)}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f>${xml(formula)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

function chartNumberReference(formula = "", values = [], formatCode = "General") {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f>${xml(formula)}</c:f><c:numCache><c:formatCode>${formatCode}</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

function chartSeriesTitleXml(item = {}, columnIndex = 1, useWorkbookRefs = false) {
  return useWorkbookRefs
    ? `<c:tx>${chartStringReference(chartCellFormula(columnIndex, 1), [item.name])}</c:tx>`
    : `<c:tx><c:v>${xml(item.name)}</c:v></c:tx>`;
}

function chartSeriesXml(chart = {}, categories = [], series = [], options = {}) {
  const formatCode = xml(chart.valueFormat || chart.numberFormat || "General");
  const catPoints = categories.map((category, index) => `<c:pt idx="${index}"><c:v>${xml(category)}</c:v></c:pt>`).join("");
  const useWorkbookRefs = Boolean(options.workbookRefs);
  const categoryXml = useWorkbookRefs && categories.length
    ? chartStringReference(chartRangeFormula(0, 2, 0, categories.length + 1), categories)
    : `<c:strLit><c:ptCount val="${categories.length}"/>${catPoints}</c:strLit>`;
  return series.map((item, index) => {
    const valPoints = item.values.map((value, valueIndex) => `<c:pt idx="${valueIndex}"><c:v>${value}</c:v></c:pt>`).join("");
    const sourceIndex = Number.isFinite(Number(item.sourceIndex)) ? Number(item.sourceIndex) : index;
    const valueColumnIndex = sourceIndex + 1;
    const titleXml = chartSeriesTitleXml(item, valueColumnIndex, useWorkbookRefs);
    const valueXml = useWorkbookRefs && item.values.length
      ? chartNumberReference(chartRangeFormula(valueColumnIndex, 2, valueColumnIndex, item.values.length + 1), item.values, formatCode)
      : `<c:numLit><c:formatCode>${formatCode}</c:formatCode><c:ptCount val="${item.values.length}"/>${valPoints}</c:numLit>`;
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${titleXml}${chartSeriesShape(chart, sourceIndex)}<c:cat>${categoryXml}</c:cat><c:val>${valueXml}</c:val></c:ser>`;
  }).join("");
}

function comboSeriesGroups(chart = {}, categories = [], seriesItems = []) {
  const requestedLineNames = new Set((Array.isArray(chart.lineSeries) ? chart.lineSeries : []).map((name) => String(name).toLowerCase()));
  const splitAt = Number.isFinite(Number(chart.lineSeriesStartIndex)) ? Math.max(0, Number(chart.lineSeriesStartIndex)) : Math.max(0, seriesItems.length - 1);
  const columnSeries = [];
  const lineSeries = [];
  for (const [index, item] of seriesItems.entries()) {
    const kind = String(item.chartType || item.type || "").toLowerCase();
    const wantsLine = kind === "line" || requestedLineNames.has(String(item.name).toLowerCase()) || (!kind && index >= splitAt);
    (wantsLine ? lineSeries : columnSeries).push(item);
  }
  return {
    columnSeries: columnSeries.length ? columnSeries : seriesItems.slice(0, Math.max(1, seriesItems.length - 1)),
    lineSeries: lineSeries.length ? lineSeries : seriesItems.slice(-1),
  };
}

function normalizeScatterSeries(chart = {}) {
  if (Array.isArray(chart.series) && chart.series.length) {
    return chart.series.slice(0, 12).map((item, index) => {
      const rawPoints = Array.isArray(item?.points) ? item.points : Array.isArray(item?.values) ? item.values : [];
      const points = rawPoints.map((point, pointIndex) => {
        if (Array.isArray(point)) return { x: Number(point[0]) || 0, y: Number(point[1]) || 0 };
        if (point && typeof point === "object") return { x: Number(point.x ?? point[0] ?? pointIndex + 1) || 0, y: Number(point.y ?? point.value ?? point[1]) || 0 };
        return { x: pointIndex + 1, y: Number(point) || 0 };
      }).slice(0, 100);
      return { name: item?.name || item?.seriesName || `Series ${index + 1}`, sourceIndex: index, points };
    });
  }
  const rows = Array.isArray(chart.points) ? chart.points : Array.isArray(chart.values) ? chart.values : [];
  const points = rows.map((row, index) => {
    if (Array.isArray(row)) return { x: Number(row[0]) || 0, y: Number(row[1]) || 0 };
    if (row && typeof row === "object") return { x: Number(row.x ?? row[0] ?? index + 1) || 0, y: Number(row.y ?? row.value ?? row[1]) || 0 };
    return { x: index + 1, y: Number(row) || 0 };
  }).slice(0, 100);
  return [{ name: chart.seriesName || chart.title || "Series 1", sourceIndex: 0, points }];
}

function scatterSeriesXml(chart = {}, series = [], options = {}) {
  const formatCode = xml(chart.valueFormat || chart.numberFormat || "General");
  const xFormatCode = xml(chart.xValueFormat || chart.xNumberFormat || formatCode);
  const yFormatCode = xml(chart.yValueFormat || chart.yNumberFormat || formatCode);
  const useWorkbookRefs = Boolean(options.workbookRefs);
  return series.map((item, index) => {
    const xPoints = item.points.map((point, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${point.x}</c:v></c:pt>`).join("");
    const yPoints = item.points.map((point, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${point.y}</c:v></c:pt>`).join("");
    const xValues = item.points.map((point) => point.x);
    const yValues = item.points.map((point) => point.y);
    const sourceIndex = Number.isFinite(Number(item.sourceIndex)) ? Number(item.sourceIndex) : index;
    const xColumnIndex = sourceIndex * 2;
    const yColumnIndex = xColumnIndex + 1;
    const titleXml = useWorkbookRefs
      ? `<c:tx>${chartStringReference(chartCellFormula(xColumnIndex, 1), [item.name])}</c:tx>`
      : `<c:tx><c:v>${xml(item.name)}</c:v></c:tx>`;
    const xXml = useWorkbookRefs && item.points.length
      ? chartNumberReference(chartRangeFormula(xColumnIndex, 3, xColumnIndex, item.points.length + 2), xValues, xFormatCode)
      : `<c:numLit><c:formatCode>${xFormatCode}</c:formatCode><c:ptCount val="${item.points.length}"/>${xPoints}</c:numLit>`;
    const yXml = useWorkbookRefs && item.points.length
      ? chartNumberReference(chartRangeFormula(yColumnIndex, 3, yColumnIndex, item.points.length + 2), yValues, yFormatCode)
      : `<c:numLit><c:formatCode>${yFormatCode}</c:formatCode><c:ptCount val="${item.points.length}"/>${yPoints}</c:numLit>`;
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${titleXml}${chartSeriesShape(chart, sourceIndex)}<c:xVal>${xXml}</c:xVal><c:yVal>${yXml}</c:yVal></c:ser>`;
  }).join("");
}

function chartWorkbookRows(chart = {}) {
  const tag = chartTypeTag(chart.chartType);
  if (tag === "c:scatterChart") {
    const series = normalizeScatterSeries(chart);
    const maxPoints = Math.max(0, ...series.map((item) => item.points.length));
    const rows = [
      series.flatMap((item) => [item.name, ""]),
      series.flatMap(() => ["X", "Y"]),
    ];
    for (let pointIndex = 0; pointIndex < maxPoints; pointIndex += 1) {
      rows.push(series.flatMap((item) => {
        const point = item.points[pointIndex];
        return point ? [point.x, point.y] : ["", ""];
      }));
    }
    return rows;
  }
  const normalized = normalizeChartSeries(chart);
  return [
    ["Category", ...normalized.series.map((item) => item.name)],
    ...normalized.categories.map((category, index) => [category, ...normalized.series.map((item) => item.values[index] ?? 0)]),
  ];
}

function embeddedChartWorkbook(chart = {}) {
  return createGeneratedXlsx({ sheets: [{ name: "Chart Data", rows: chartWorkbookRows(chart) }] });
}

function chartRelsXml(workbookName = "chartData1.xlsx") {
  return rels([{ id: "rWorkbook", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", target: `../embeddings/${workbookName}` }]);
}

function chartXml(chart = {}, options = {}) {
  const useExternalWorkbook = Boolean(options.externalWorkbook);
  const normalized = normalizeChartSeries(chart);
  const categories = normalized.categories;
  const seriesItems = normalized.series;
  const chartTag = chartTypeTag(chart.chartType);
  const isCombo = chartTag === "combo";
  const isBar = chartTag === "c:barChart";
  const isPieLike = chartTag === "c:pieChart" || chartTag === "c:doughnutChart";
  const isScatter = chartTag === "c:scatterChart";
  const title = chart.title ? `<c:title>${chartRichText(chart.title)}<c:layout/></c:title>` : "";
  const formatCode = xml(chart.valueFormat || chart.numberFormat || "General");
  const comboGroups = isCombo ? comboSeriesGroups(chart, categories, seriesItems) : null;
  const series = isScatter ? `${scatterSeriesXml(chart, normalizeScatterSeries(chart), { workbookRefs: useExternalWorkbook })}${chartDataLabels(chart)}` : `${chartSeriesXml(chart, categories, seriesItems, { workbookRefs: useExternalWorkbook })}${chartDataLabels(chart)}`;
  const grouping = isScatter ? `<c:scatterStyle val="${chart.scatterStyle || "marker"}"/>` : chartTag === "c:lineChart" || chartTag === "c:areaChart" ? `<c:grouping val="standard"/>` : isBar ? `<c:barDir val="${chart.barDirection === "horizontal" ? "bar" : "col"}"/><c:grouping val="clustered"/>` : chartTag === "c:doughnutChart" ? `<c:holeSize val="${Math.min(90, Math.max(10, Number(chart.holeSize || 50)))}"/>` : "";
  const axes = isPieLike ? "" : `<c:axId val="123456"/><c:axId val="654321"/>`;
  const axisXml = isPieLike ? "" : isScatter ? `<c:valAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>${chartAxisTitle(chart.categoryAxisTitle || chart.xAxisTitle)}<c:numFmt formatCode="${xml(chart.xValueFormat || chart.xNumberFormat || formatCode)}" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="654321"/><c:crosses val="autoZero"/></c:valAx><c:valAx><c:axId val="654321"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>${chart.showGridLines === false ? "" : "<c:majorGridlines/>"}${chartAxisTitle(chart.valueAxisTitle || chart.yAxisTitle)}<c:numFmt formatCode="${xml(chart.yValueFormat || chart.yNumberFormat || formatCode)}" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="123456"/><c:crosses val="autoZero"/></c:valAx>` : `<c:catAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>${chartAxisTitle(chart.categoryAxisTitle || chart.xAxisTitle)}<c:tickLblPos val="nextTo"/><c:crossAx val="654321"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="654321"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>${chart.showGridLines === false ? "" : "<c:majorGridlines/>"}${chartAxisTitle(chart.valueAxisTitle || chart.yAxisTitle)}<c:numFmt formatCode="${formatCode}" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="123456"/><c:crosses val="autoZero"/></c:valAx>`;
  const chartBody = isCombo
    ? `<c:barChart><c:barDir val="${chart.barDirection === "horizontal" ? "bar" : "col"}"/><c:grouping val="clustered"/>${chartSeriesXml(chart, categories, comboGroups.columnSeries, { workbookRefs: useExternalWorkbook })}${axes}</c:barChart><c:lineChart><c:grouping val="standard"/>${chartSeriesXml(chart, categories, comboGroups.lineSeries, { workbookRefs: useExternalWorkbook })}${axes}</c:lineChart>`
    : "<" + chartTag + ">" + grouping + series + axes + "</" + chartTag + ">";
  const externalData = useExternalWorkbook ? `<c:externalData r:id="rWorkbook"><c:autoUpdate val="0"/></c:externalData>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/>${externalData}<c:chart>${title}<c:plotArea><c:layout/>${chartBody}${axisXml}</c:plotArea><c:legend><c:legendPos val="${chartLegendPosition(chart.legendPosition)}"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

function generatedChartFrame(id, relId, chart = {}, index = 0) {
  const x = emu(chart.left, 70);
  const y = emu(chart.top, 170 + index * 20);
  const cx = emu(chart.width, 560);
  const cy = emu(chart.height, 260);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${xml(chart.title || `Chart ${index + 1}`)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

const MAX_EMBEDDED_IMAGE_BYTES = Number(process.env.GENERATED_OFFICE_MAX_IMAGE_BYTES || 8 * 1024 * 1024);
const EMBEDDED_IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"]);

function embeddedImageBytes(image = {}) {
  if (!image?.base64 || typeof image.base64 !== "string") return null;
  const type = String(image.type || "image/png").split(";", 1)[0].trim().toLowerCase();
  if (!EMBEDDED_IMAGE_TYPES.has(type)) throw new Error(`Unsupported embedded image type: ${type || "unknown"}.`);
  const encoded = image.base64.replace(/^data:[^,]+,/, "");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(encoded) || encoded.length % 4 === 1) throw new Error("Embedded image content is not valid base64.");
  const bytes = Buffer.from(encoded, "base64");
  if (!bytes.length || bytes.length > MAX_EMBEDDED_IMAGE_BYTES) throw new Error("Embedded image exceeds the configured size limit.");
  const signature = type === "image/png"
    ? bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))
    : type === "image/jpeg"
      ? bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff
      : type === "image/gif"
        ? bytes.length >= 6 && ["GIF87a", "GIF89a"].includes(bytes.subarray(0, 6).toString("ascii"))
        : type === "image/bmp"
          ? bytes.length >= 2 && bytes.subarray(0, 2).toString("ascii") === "BM"
          : bytes.length >= 12 && bytes.subarray(0, 4).toString("ascii") === "RIFF" && bytes.subarray(8, 12).toString("ascii") === "WEBP";
  if (!signature) throw new Error("Embedded image content does not match its declared type.");
  return bytes;
}
function imageExt(type = "image/png") {
  if (type.includes("jpeg") || type.includes("jpg")) return "jpg";
  if (type.includes("gif")) return "gif";
  if (type.includes("webp")) return "webp";
  if (type.includes("bmp")) return "bmp";
  return "png";
}

function imageContentType(ext) {
  return ext === "jpg" ? "image/jpeg" : `image/${ext}`;
}

function ratioFromImage(image = {}) {
  const width = asNumber(image.pixelWidth ?? image.naturalWidth ?? image.sourceWidth, 0);
  const height = asNumber(image.pixelHeight ?? image.naturalHeight ?? image.sourceHeight, 0);
  return width > 0 && height > 0 ? width / height : 0;
}

function pct100k(value, fallback = 0) {
  const numeric = asNumber(value, fallback);
  if (!Number.isFinite(numeric)) return 0;
  if (Math.abs(numeric) <= 1) return Math.round(Math.max(0, Math.min(1, numeric)) * 100000);
  return Math.round(Math.max(0, Math.min(100, numeric)) * 1000);
}

function pictureCropXml(image = {}, frameRatio = 0) {
  const crop = image.crop && typeof image.crop === "object" ? image.crop : {};
  const explicit = ["left", "right", "top", "bottom"].some((key) => crop[key] !== undefined || image[`crop${key[0].toUpperCase()}${key.slice(1)}`] !== undefined);
  if (explicit) {
    const l = pct100k(crop.left ?? image.cropLeft);
    const r = pct100k(crop.right ?? image.cropRight);
    const t = pct100k(crop.top ?? image.cropTop);
    const b = pct100k(crop.bottom ?? image.cropBottom);
    return `<a:srcRect l="${l}" r="${r}" t="${t}" b="${b}"/>`;
  }
  const mode = String(image.fit || image.sizing || image.mode || "stretch").toLowerCase();
  const sourceRatio = ratioFromImage(image);
  if (!(mode === "fill" || mode === "cover") || !sourceRatio || !frameRatio) return "";
  if (sourceRatio > frameRatio) {
    const cropEachSide = Math.round(((sourceRatio - frameRatio) / sourceRatio / 2) * 100000);
    return `<a:srcRect l="${cropEachSide}" r="${cropEachSide}"/>`;
  }
  if (sourceRatio < frameRatio) {
    const cropTopBottom = Math.round(((frameRatio - sourceRatio) / frameRatio / 2) * 100000);
    return `<a:srcRect t="${cropTopBottom}" b="${cropTopBottom}"/>`;
  }
  return "";
}

function pictureFrame(image = {}, index = 0) {
  const left = asNumber(image.left, 390);
  const top = asNumber(image.top, 130 + index * 170);
  let width = asNumber(image.width, 260);
  let height = asNumber(image.height, 150);
  let x = left;
  let y = top;
  const mode = String(image.fit || image.sizing || image.mode || "stretch").toLowerCase();
  const sourceRatio = ratioFromImage(image);
  const frameRatio = width > 0 && height > 0 ? width / height : 0;
  if ((mode === "fit" || mode === "contain") && sourceRatio && frameRatio) {
    if (sourceRatio > frameRatio) {
      const fittedHeight = width / sourceRatio;
      y += (height - fittedHeight) / 2;
      height = fittedHeight;
    } else if (sourceRatio < frameRatio) {
      const fittedWidth = height * sourceRatio;
      x += (width - fittedWidth) / 2;
      width = fittedWidth;
    }
  }
  return { x: emu(x, 390), y: emu(y, 130 + index * 170), cx: emu(width, 260), cy: emu(height, 150), frameRatio };
}

function picture(id, relId, image, index) {
  const { x, y, cx, cy, frameRatio } = pictureFrame(image, index);
  const name = xml(image.altText || image.name || `Image ${index + 1}`);
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}" descr="${name}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relId}"/>${pictureCropXml(image, frameRatio)}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function imageBytes(image) {
  if (!image?.base64) return null;
  return embeddedImageBytes(image);
}

function normalizeGeneratedImages(value) {
  return Array.isArray(value) ? value.filter((image) => image && typeof image === "object" && image.base64).slice(0, 50) : [];
}

async function resolveImageUrls(images = []) {
  const resolved = [];
  for (const image of Array.isArray(images) ? images : []) {
    if (!image || typeof image !== "object") continue;
    if (image.base64 || !image.imageUrl) {
      resolved.push(image);
      continue;
    }
    const asset = await fetchImageAsset(String(image.imageUrl));
    resolved.push({ ...image, name: image.name || asset.name, type: image.type || asset.type, pixelWidth: image.pixelWidth || asset.pixelWidth, pixelHeight: image.pixelHeight || asset.pixelHeight, base64: asset.base64, altText: image.altText || asset.name });
  }
  return resolved;
}

async function resolveGeneratedPayloadAssets(payload = {}, kind = "") {
  if (!payload || typeof payload !== "object") return payload;
  if (kind === "pptx") {
    const slides = Array.isArray(payload.slides) ? [] : payload.slides;
    if (Array.isArray(payload.slides)) {
      for (const slide of payload.slides) slides.push(slide && typeof slide === "object" ? { ...slide, images: await resolveImageUrls(slide.images) } : slide);
      return { ...payload, slides };
    }
  }
  if (kind === "docx") {
    const sections = Array.isArray(payload.sections) ? [] : payload.sections;
    if (Array.isArray(payload.sections)) {
      for (const section of payload.sections) sections.push(section && typeof section === "object" ? { ...section, images: await resolveImageUrls(section.images) } : section);
    }
    return { ...payload, images: await resolveImageUrls(payload.images), ...(Array.isArray(payload.sections) ? { sections } : {}) };
  }
  if (kind === "xlsx") {
    const sheets = Array.isArray(payload.sheets) ? [] : payload.sheets;
    if (Array.isArray(payload.sheets)) {
      for (const sheet of payload.sheets) sheets.push(sheet && typeof sheet === "object" ? { ...sheet, images: await resolveImageUrls(sheet.images) } : sheet);
      return { ...payload, sheets };
    }
    return { ...payload, images: await resolveImageUrls(payload.images) };
  }
  return payload;
}

function slideXml(slide, slideIndex, imageRels, chartRels = [], hyperlinkRels = []) {
  const title = slide.title || `Slide ${slideIndex + 1}`;
  const body = slide.body || slide.content || "";
  let nextId = 2;
  const shapes = [
    textBox(nextId++, "Title", title, emu(42, 42), emu(28, 28), emu(650, 650), emu(58, 58), 3200, true),
    bodyBox(nextId++, "Body", body, emu(52, 52), emu(105, 105), emu(560, 560), emu(360, 360)),
    generatedSlideObjects(nextId, slide, imageRels, chartRels, hyperlinkRels),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${slideBackground(slide)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function templateSlideXml(slide, slideIndex, layoutXml = "", imageRels = [], chartRels = [], hyperlinkRels = []) {
  const placeholders = placeholderShapesFromLayout(layoutXml);
  if (!placeholders.length) return slideXml(slide, slideIndex, imageRels, chartRels, hyperlinkRels);
  const used = new Set();
  const title = slide.title || `Slide ${slideIndex + 1}`;
  const body = slide.body || slide.content || "";
  const subtitle = slide.subtitle || slide.kicker || "";
  const renderedPlaceholders = placeholders.map((placeholder) => {
    let text = "";
    if (placeholder.kind === "title") text = title;
    else if (placeholder.kind === "subtitle") text = subtitle || body;
    else if (placeholder.kind === "body") text = body;
    else text = "";
    if (text) used.add(placeholder.kind);
    return replaceShapeText(placeholder.xml, text);
  }).join("");
  const highestPlaceholderId = Math.max(501, ...placeholders.map((placeholder) => placeholder.id || 0));
  const generatedTitleId = highestPlaceholderId + 1;
  const generatedBodyId = highestPlaceholderId + 2;
  const fallbackTitle = used.has("title") ? "" : textBox(generatedTitleId, "Generated Title", title, emu(42, 42), emu(28, 28), emu(650, 650), emu(58, 58), 3200, true);
  const fallbackBody = body && !used.has("body") && !used.has("subtitle") ? bodyBox(generatedBodyId, "Generated Body", body, emu(52, 52), emu(105, 105), emu(560, 560), emu(360, 360)) : "";
  const extraObjects = generatedSlideObjects(highestPlaceholderId + 3, slide, imageRels, chartRels, hyperlinkRels);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld>${slideBackground(slide)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${renderedPlaceholders}${fallbackTitle}${fallbackBody}${extraObjects}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function notesXml(notes = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${bodyBox(2, "Notes", notes, 0, 0, SLIDE_W, SLIDE_H)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

function rels(rels) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((rel) => `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${xml(rel.target)}"${rel.targetMode ? ` TargetMode="${xml(rel.targetMode)}"` : ""}/>`).join("")}</Relationships>`;
}

function contentTypes(slideCount, noteCount, imageExts, chartCount = 0) {
  const defaults = new Set(["xml", "rels", ...imageExts]);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[...defaults].map((ext) => `<Default Extension="${ext}" ContentType="${ext === "rels" ? "application/vnd.openxmlformats-package.relationships+xml" : ext === "xml" ? "application/xml" : imageContentType(ext)}"/>`).join("")}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${Array.from({ length: slideCount }, (_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("")}${Array.from({ length: noteCount }, (_, i) => `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`).join("")}${Array.from({ length: chartCount }, (_, i) => `<Override PartName="/ppt/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}</Types>`;
}

const ROOT_RELS = rels([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" }]);
function generatedTheme(brand = DEFAULT_BRAND) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CTRL Generated"><a:themeElements><a:clrScheme name="CTRL"><a:dk1><a:srgbClr val="${brand.dark}"/></a:dk1><a:lt1><a:srgbClr val="${brand.light}"/></a:lt1><a:accent1><a:srgbClr val="${brand.accent1}"/></a:accent1><a:accent2><a:srgbClr val="${brand.accent2}"/></a:accent2><a:accent3><a:srgbClr val="${brand.accent3}"/></a:accent3></a:clrScheme><a:fontScheme name="CTRL"><a:majorFont><a:latin typeface="${xml(brand.majorFont)}"/></a:majorFont><a:minorFont><a:latin typeface="${xml(brand.minorFont)}"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`;
}
const SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles/></p:sldMaster>`;
const SLIDE_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

export function createGeneratedPptx(payload = {}) {
  const templatePreserved = createTemplatePreservingPptx(payload);
  if (templatePreserved) return templatePreserved;

  const rawSlides = Array.isArray(payload.slides) && payload.slides.length ? payload.slides.slice(0, 100) : [{ title: payload.title || "Generated deck", body: "" }];
  const slides = rawSlides.map((slide, index) => deckChromeSlide(slide, payload, index, rawSlides.length));
  const brand = normalizedBrandProfile(payload);
  const entries = [];
  const imageExts = new Set();
  let mediaIndex = 1;
  let noteCount = 0;
  let chartIndex = 1;
  let embeddedWorkbookIndex = 1;

  const presentationRels = [
    { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "slideMasters/slideMaster1.xml" },
    { id: "rTheme", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "theme/theme1.xml" },
  ];

  entries.push(["_rels/.rels", ROOT_RELS]);
  entries.push(["ppt/theme/theme1.xml", generatedTheme(brand)]);
  entries.push(["ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER]);
  entries.push(["ppt/slideMasters/_rels/slideMaster1.xml.rels", rels([
    { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: "../slideLayouts/slideLayout1.xml" },
    { id: "rId2", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "../theme/theme1.xml" },
  ])]);
  entries.push(["ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT]);
  entries.push(["ppt/slideLayouts/_rels/slideLayout1.xml.rels", rels([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "../slideMasters/slideMaster1.xml" }])]);

  for (const [index, slide] of slides.entries()) {
    const slideNumber = index + 1;
    const slideRels = [{ id: "rLayout", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: "../slideLayouts/slideLayout1.xml" }];
    const imageRels = [];
    const chartRels = [];
    const hyperlinkRels = [];
    const images = Array.isArray(slide.images) ? slide.images.slice(0, 8) : [];
    for (const image of images) {
      if (!image?.base64) continue;
      const ext = imageExt(image.type);
      imageExts.add(ext);
      const mediaName = `image${mediaIndex++}.${ext}`;
      const relId = `rImg${mediaIndex}`;
      entries.push([`ppt/media/${mediaName}`, embeddedImageBytes(image)]);
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `../media/${mediaName}` });
      imageRels.push({ relId, image: { ...image, name: image.name || mediaName } });
    }
    const charts = Array.isArray(slide.charts) ? slide.charts.slice(0, 6) : [];
    for (const chart of charts) {
      if (!chart || typeof chart !== "object") continue;
      const styledChart = brandChart(chart, slide.__brand);
      const chartNumber = chartIndex++;
      const workbookName = `chartData${embeddedWorkbookIndex++}.xlsx`;
      const relId = `rChart${chartNumber}`;
      entries.push([`ppt/charts/chart${chartNumber}.xml`, chartXml(styledChart, { externalWorkbook: true })]);
      entries.push([`ppt/charts/_rels/chart${chartNumber}.xml.rels`, chartRelsXml(workbookName)]);
      entries.push([`ppt/embeddings/${workbookName}`, embeddedChartWorkbook(styledChart)]);
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: `../charts/chart${chartNumber}.xml` });
      chartRels.push({ relId, chart: styledChart });
    }
    const links = Array.isArray(slide.links) ? slide.links.slice(0, 12) : [];
    for (const link of links) {
      if (!link || typeof link !== "object" || !link.url) continue;
      const relId = `rLink${hyperlinkRels.length + 1}`;
      slideRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: String(link.url), targetMode: "External" });
      hyperlinkRels.push({ relId, link });
    }
    if (slide.notes) {
      noteCount += 1;
      entries.push([`ppt/notesSlides/notesSlide${slideNumber}.xml`, notesXml(slide.notes)]);
      slideRels.push({ id: "rNotes", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide", target: `../notesSlides/notesSlide${slideNumber}.xml` });
    }
    entries.push([`ppt/slides/slide${slideNumber}.xml`, slideXml(slide, index, imageRels, chartRels, hyperlinkRels)]);
    entries.push([`ppt/slides/_rels/slide${slideNumber}.xml.rels`, rels(slideRels)]);
    presentationRels.push({ id: `rSlide${slideNumber}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: `slides/slide${slideNumber}.xml` });
  }

  const sldIdLst = slides.map((_slide, index) => `<p:sldId id="${256 + index}" r:id="rSlide${index + 1}"/>`).join("");
  const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIdLst}</p:sldIdLst><p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}" type="wide"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  entries.push(["ppt/presentation.xml", presentation]);
  entries.push(["ppt/_rels/presentation.xml.rels", rels(presentationRels)]);
  const types = contentTypes(slides.length, noteCount, imageExts, chartIndex - 1).replace("</Types>", `${Array.from({ length: embeddedWorkbookIndex - 1 }, (_unused, index) => `<Override PartName="/ppt/embeddings/chartData${index + 1}.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`).join("")}</Types>`);
  entries.unshift(["[Content_Types].xml", types]);
  return zip(entries);
}


function wordParagraph(text = "", styleId = "") {
  const style = styleId ? `<w:pPr><w:pStyle w:val="${xml(styleId)}"/></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`;
}

function safeBookmarkName(value = "") {
  const cleaned = String(value || "").replace(/[^A-Za-z0-9_]/g, "_").replace(/^([^A-Za-z_])/, "_$1").slice(0, 40);
  return cleaned || "CTRL_REF";
}

function wordCaptionParagraph(caption = {}, bookmarkId = 1, sequenceNumber = 1) {
  const label = xml(caption.label || caption.type || "Figure");
  const text = xml(caption.text || caption.caption || "");
  const bookmark = safeBookmarkName(caption.bookmark || caption.id || `${label}_${bookmarkId}`);
  const suffix = text ? `: ${text}` : "";
  return `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:bookmarkStart w:id="${bookmarkId}" w:name="${xml(bookmark)}"/><w:r><w:t xml:space="preserve">${label} </w:t></w:r><w:fldSimple w:instr="SEQ ${label} \\* ARABIC"><w:r><w:t>${sequenceNumber}</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve">${suffix}</w:t></w:r><w:bookmarkEnd w:id="${bookmarkId}"/></w:p>`;
}

function wordCrossReferenceParagraph(reference = {}) {
  const bookmark = safeBookmarkName(reference.bookmark || reference.id || reference.target || "");
  const prefix = reference.text || reference.label || "See";
  const fallback = reference.fallback || reference.caption || bookmark;
  return `<w:p><w:r><w:t xml:space="preserve">${xml(prefix)} </w:t></w:r><w:fldSimple w:instr="REF ${xml(bookmark)} \\h"><w:r><w:t xml:space="preserve">${xml(fallback)}</w:t></w:r></w:fldSimple></w:p>`;
}

function wordListParagraph(text = "", numId = 1, level = 0) {
  const ilvl = Math.max(0, Math.min(8, Math.round(asNumber(level, 0))));
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`;
}

function wordListItems(items = [], type = "bullet") {
  if (!Array.isArray(items)) return "";
  const numId = type === "number" || type === "numbered" || type === "ordered" ? 2 : 1;
  return items.map((item) => {
    if (item && typeof item === "object") return wordListParagraph(item.text || item.value || "", numId, item.level || 0);
    return wordListParagraph(item, numId, 0);
  }).join("");
}

function wordHyperlinkParagraph(relId, link = {}) {
  const text = link.text || link.label || link.url || "Link";
  return `<w:p><w:hyperlink r:id="${relId}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:u w:val="single"/><w:color w:val="0563C1"/></w:rPr><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:hyperlink></w:p>`;
}

function wordTableOfContents() {
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Table of contents</w:t></w:r></w:p><w:p><w:fldSimple w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u"><w:r><w:t>Right-click and update field in Word to refresh this table of contents.</w:t></w:r></w:fldSimple></w:p>`;
}

function wordFootnoteReference(id) {
  return `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
}

function wordEndnoteReference(id) {
  return `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;
}

function wordParagraphWithEndnote(text = "", endnoteId = null, styleId = "") {
  const style = styleId ? `<w:pPr><w:pStyle w:val="${xml(styleId)}"/></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r>${endnoteId ? wordEndnoteReference(endnoteId) : ""}</w:p>`;
}

function wordRevisionRun(revision = {}, index = 0) {
  const type = revision.type === "delete" || revision.type === "deleted" || revision.type === "deletion" ? "del" : "ins";
  const tag = type === "del" ? "w:del" : "w:ins";
  const runTag = type === "del" ? "w:delText" : "w:t";
  const author = xml(revision.author || "CTRL AI");
  const date = xml(revision.date || new Date(0).toISOString());
  const text = xml(revision.text || revision.value || "");
  return `<${tag} w:id="${index + 1}" w:author="${author}" w:date="${date}"><w:r><${runTag} xml:space="preserve">${text}</${runTag}></w:r></${tag}>`;
}

function wordRevisionParagraph(revisions = []) {
  const normalized = Array.isArray(revisions) ? revisions.filter((revision) => revision && typeof revision === "object" && (revision.text || revision.value)) : [];
  if (!normalized.length) return "";
  return `<w:p>${normalized.map((revision, index) => wordRevisionRun(revision, index)).join("")}</w:p>`;
}

function wordParagraphWithFootnote(text = "", footnoteId = null, styleId = "") {
  const style = styleId ? `<w:pPr><w:pStyle w:val="${xml(styleId)}"/></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r>${footnoteId ? wordFootnoteReference(footnoteId) : ""}</w:p>`;
}

function wordTable(rows = []) {
  const normalized = rows.map((row) => Array.isArray(row) ? row : [row]);
  if (!normalized.length) return "";
  const width = Math.floor(9000 / Math.max(1, ...normalized.map((row) => row.length)));
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>${Array.from({ length: Math.max(1, ...normalized.map((row) => row.length)) }, () => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${normalized.map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${wordParagraph(String(cell ?? ""))}</w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
}

function wordImageParagraph(relId, image = {}, index = 1) {
  const cx = emu(image.width, 420);
  const cy = emu(image.height, 240);
  const name = xml(image.altText || image.name || `Image ${index}`);
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${index}" name="${name}" descr="${name}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${index}" name="${name}" descr="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function docxStyles() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="44"/><w:color w:val="1F4E79"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2563EB"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="374151"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`;
}


function docxHeader(text = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${wordParagraph(text)}</w:hdr>`;
}

function docxFooter(text = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${wordParagraph(text)}</w:ftr>`;
}

function docxComments(comments = []) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments.map((comment, index) => `<w:comment w:id="${index}" w:author="${xml(comment.author || "CTRL AI")}" w:date="${xml(comment.date || new Date(0).toISOString())}">${wordParagraph(comment.text || comment)}</w:comment>`).join("")}</w:comments>`;
}

function docxFootnotes(footnotes = []) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${footnotes.map((footnote, index) => `<w:footnote w:id="${index + 1}">${wordParagraph(footnote.text || footnote)}</w:footnote>`).join("")}</w:footnotes>`;
}

function docxEndnotes(endnotes = []) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>${endnotes.map((endnote, index) => `<w:endnote w:id="${index + 1}">${wordParagraph(endnote.text || endnote)}</w:endnote>`).join("")}</w:endnotes>`;
}

function docxNumbering() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="?"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>`;
}

function docxSettings(trackRevisions = false) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${trackRevisions ? "<w:trackRevisions/>" : ""}</w:settings>`;
}

function normalizedColumns(columns) {
  if (typeof columns === "number") return { count: columns };
  return columns && typeof columns === "object" ? columns : {};
}

function docxColumns(payload = {}) {
  const columns = normalizedColumns(payload.columns || payload.page?.columns);
  const count = Math.max(1, Math.min(8, Math.round(asNumber(columns.count || columns.num, 1))));
  if (count <= 1 && !columns.separator && !columns.space) return "";
  const space = Math.max(0, Math.round(asNumber(columns.space, 720)));
  return `<w:cols w:num="${count}" w:space="${space}"${columns.separator ? ' w:sep="1"' : ""}/>`;
}

function sectionLayoutPayload(defaultPayload = {}, section = {}) {
  return {
    ...defaultPayload,
    page: { ...(defaultPayload.page || {}), ...(section.page || {}) },
    columns: section.columns ?? section.page?.columns ?? defaultPayload.columns ?? defaultPayload.page?.columns,
  };
}

function hasSectionLayout(section = {}) {
  return Boolean(section && typeof section === "object" && (section.page || section.columns || section.page?.columns));
}

function wordSectionBreak(payload = {}, sectRefs = "") {
  return `<w:p><w:pPr>${docxSectPr(payload, sectRefs)}</w:pPr></w:p>`;
}

function docxSectPr(payload = {}, sectRefs = "") {
  const layout = payload.page || {};
  const orientation = layout.orientation === "landscape" ? "landscape" : "portrait";
  const width = orientation === "landscape" ? 15840 : 12240;
  const height = orientation === "landscape" ? 12240 : 15840;
  const margins = layout.margins || {};
  const margin = (name, fallback) => Math.max(0, Math.round(asNumber(margins[name], fallback)));
  return `<w:sectPr>${sectRefs}<w:pgSz w:w="${width}" w:h="${height}"${orientation === "landscape" ? ' w:orient="landscape"' : ""}/><w:pgMar w:top="${margin("top", 1440)}" w:right="${margin("right", 1440)}" w:bottom="${margin("bottom", 1440)}" w:left="${margin("left", 1440)}"/>${docxColumns(payload)}</w:sectPr>`;
}

export function createGeneratedDocx(payload = {}) {
  const templateEntries = decodeDocxTemplate(payload.templateBase64);
  const entries = [];
  const title = payload.title || "Generated document";
  const sections = Array.isArray(payload.sections) && payload.sections.length ? payload.sections : [{ heading: title, body: payload.body || "" }];
  const comments = Array.isArray(payload.comments) ? payload.comments.slice(0, 100) : [];
  const footnotes = Array.isArray(payload.footnotes) ? payload.footnotes.slice(0, 100) : [];
  const endnotes = Array.isArray(payload.endnotes) ? payload.endnotes.slice(0, 100) : [];
  const revisions = Array.isArray(payload.revisions) ? payload.revisions.slice(0, 200) : [];
  const hasLists = Array.isArray(payload.bullets) || Array.isArray(payload.numberedList) || sections.some((section) => Array.isArray(section.bullets) || Array.isArray(section.numberedList));
  const headerText = typeof payload.header === "string" ? payload.header : "";
  const footerText = typeof payload.footer === "string" ? payload.footer : "";
  const docRels = [];
  const imageExts = new Set();
  let mediaIndex = templateEntries ? [...templateEntries.keys()].filter((name) => /^word\/media\//.test(name)).length + 1 : 1;
  let relIndex = 1;
  let bookmarkIndex = 1;
  const captionCounters = new Map();
  const bodyParts = [wordParagraph(title, "Title")];
  const addCaption = (caption) => {
    if (!caption || typeof caption !== "object") return;
    const label = caption.label || caption.type || "Figure";
    const next = (captionCounters.get(label) || 0) + 1;
    captionCounters.set(label, next);
    bodyParts.push(wordCaptionParagraph(caption, bookmarkIndex++, next));
  };
  const addCrossReference = (reference) => {
    if (!reference || typeof reference !== "object") return;
    bodyParts.push(wordCrossReferenceParagraph(reference));
  };
  const addWordHyperlink = (link) => {
    if (!link || typeof link !== "object" || !link.url) return;
    const relId = `rCtrlLink${relIndex++}`;
    docRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: String(link.url), targetMode: "External" });
    bodyParts.push(wordHyperlinkParagraph(relId, link));
  };
  if (payload.tableOfContents) bodyParts.push(wordTableOfContents());
  if (revisions.length) bodyParts.push(wordRevisionParagraph(revisions));
  if (Array.isArray(payload.bullets)) bodyParts.push(wordListItems(payload.bullets, "bullet"));
  if (Array.isArray(payload.numberedList)) bodyParts.push(wordListItems(payload.numberedList, "number"));

  const addWordImage = (image) => {
    const bytes = imageBytes(image);
    if (!bytes) return;
    const ext = imageExt(image.type);
    imageExts.add(ext);
    const mediaName = `image${mediaIndex++}.${ext}`;
    const relId = `rCtrlImage${relIndex++}`;
    entries.push([`word/media/${mediaName}`, bytes]);
    docRels.push({ id: relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `media/${mediaName}` });
    bodyParts.push(wordImageParagraph(relId, { ...image, name: image.name || mediaName }, mediaIndex));
    addCaption(image.caption);
  };

  if (Array.isArray(payload.links)) payload.links.slice(0, 100).forEach(addWordHyperlink);
  for (const image of normalizeGeneratedImages(payload.images)) addWordImage(image);
  if (Array.isArray(payload.captions)) payload.captions.slice(0, 100).forEach(addCaption);
  if (Array.isArray(payload.crossReferences)) payload.crossReferences.slice(0, 100).forEach(addCrossReference);
  for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
    const section = sections[sectionIndex];
    if (section.heading) bodyParts.push(wordParagraph(section.heading, section.level === 2 ? "Heading2" : "Heading1"));
    if (section.body) String(section.body).split(/\n\s*\n/).forEach((para) => bodyParts.push(wordParagraph(para)));
    if (Array.isArray(section.revisions)) bodyParts.push(wordRevisionParagraph(section.revisions));
    if (section.footnote) {
      const footnoteId = footnotes.length + 1;
      footnotes.push({ text: section.footnote });
      bodyParts.push(wordParagraphWithFootnote("Source note", footnoteId));
    }
    if (section.endnote) {
      const endnoteId = endnotes.length + 1;
      endnotes.push({ text: section.endnote });
      bodyParts.push(wordParagraphWithEndnote("Endnote", endnoteId));
    }
    if (Array.isArray(section.bullets)) bodyParts.push(wordListItems(section.bullets, "bullet"));
    if (Array.isArray(section.numberedList)) bodyParts.push(wordListItems(section.numberedList, "number"));
    if (Array.isArray(section.links)) section.links.slice(0, 50).forEach(addWordHyperlink);
    if (Array.isArray(section.crossReferences)) section.crossReferences.slice(0, 50).forEach(addCrossReference);
    if (Array.isArray(section.table)) {
      bodyParts.push(wordTable(section.table));
      addCaption(section.tableCaption || section.caption);
    }
    for (const image of normalizeGeneratedImages(section.images)) addWordImage(image);
    if (sectionIndex < sections.length - 1 && hasSectionLayout(section)) bodyParts.push(wordSectionBreak(sectionLayoutPayload(payload, section)));
  }
  const sectRefs = [
    headerText ? '<w:headerReference w:type="default" r:id="rCtrlHeader1"/>' : "",
    footerText ? '<w:footerReference w:type="default" r:id="rCtrlFooter1"/>' : "",
    !headerText && templateEntries ? templateSectionReferences(templateEntries).match(/<w:headerReference\b[^>]*\/>/g)?.join("") || "" : "",
    !footerText && templateEntries ? templateSectionReferences(templateEntries).match(/<w:footerReference\b[^>]*\/>/g)?.join("") || "" : "",
  ].join("");
  const lastSection = sections[sections.length - 1] || {};
  const finalLayout = hasSectionLayout(lastSection) ? sectionLayoutPayload(payload, lastSection) : payload;
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyParts.join("")}${docxSectPr(finalLayout, sectRefs)}</w:body></w:document>`;
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    hasLists ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : "",
    revisions.length ? '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' : "",
    headerText ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : "",
    footerText ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : "",
    comments.length ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : "",
    footnotes.length ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : "",
    endnotes.length ? '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' : "",
  ].filter(Boolean).join("");
  const imageDefaults = [...imageExts].map((ext) => `<Default Extension="${ext}" ContentType="${imageContentType(ext)}"/>`).join("");
  entries.push(["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imageDefaults}${overrides}</Types>`]);
  entries.push(["_rels/.rels", rels([{ id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "word/document.xml" }])]);
  if (headerText) docRels.push({ id: "rCtrlHeader1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", target: "header1.xml" });
  if (footerText) docRels.push({ id: "rCtrlFooter1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", target: "footer1.xml" });
  if (comments.length) docRels.push({ id: "rCtrlComments", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", target: "comments.xml" });
  if (hasLists) docRels.push({ id: "rCtrlNumbering", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", target: "numbering.xml" });
  if (footnotes.length) docRels.push({ id: "rCtrlFootnotes", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", target: "footnotes.xml" });
  if (endnotes.length) docRels.push({ id: "rCtrlEndnotes", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", target: "endnotes.xml" });
  if (revisions.length) docRels.push({ id: "rCtrlSettings", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", target: "settings.xml" });
  entries.push(["word/document.xml", document]);
  entries.push(["word/styles.xml", docxStyles()]);
  if (hasLists) entries.push(["word/numbering.xml", docxNumbering()]);
  if (revisions.length) entries.push(["word/settings.xml", docxSettings(true)]);
  if (headerText) entries.push(["word/header1.xml", docxHeader(headerText)]);
  if (footerText) entries.push(["word/footer1.xml", docxFooter(footerText)]);
  if (comments.length) entries.push(["word/comments.xml", docxComments(comments)]);
  if (footnotes.length) entries.push(["word/footnotes.xml", docxFootnotes(footnotes)]);
  if (endnotes.length) entries.push(["word/endnotes.xml", docxEndnotes(endnotes)]);
  entries.push(["word/_rels/document.xml.rels", rels(docRels)]);
  if (!templateEntries) return zip(entries);

  const merged = new Map(templateEntries);
  for (const [name, content] of entries) merged.set(name, content);
  if (templateEntries.has("word/styles.xml")) merged.set("word/styles.xml", templateEntries.get("word/styles.xml"));
  if (templateEntries.has("word/numbering.xml") && !hasLists) merged.set("word/numbering.xml", templateEntries.get("word/numbering.xml"));
  const templateRels = templateEntries.get("word/_rels/document.xml.rels")?.toString("utf8") || "";
  merged.set("word/_rels/document.xml.rels", mergeDocxRelationships(templateRels, rels(docRels)));
  merged.set("[Content_Types].xml", mergeDocxContentTypes(templateEntries.get("[Content_Types].xml")?.toString("utf8"), entries.find(([name]) => name === "[Content_Types].xml")?.[1]?.toString?.("utf8")));
  return zipFromMap(merged);
}

function columnName(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function cellXml(value, rowIndex, colIndex) {
  const ref = `${columnName(colIndex)}${rowIndex + 1}`;
  if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
  if (typeof value === "string" && value.startsWith("=")) return `<c r="${ref}"><f>${xml(value.slice(1))}</f></c>`;
  return `<c r="${ref}" t="inlineStr"><is><t>${xml(value ?? "")}</t></is></c>`;
}

function xlsxRowHeightAttrs(options = {}, rowIndex = 0) {
  const rowNumber = rowIndex + 1;
  const heights = Array.isArray(options.rowHeights) ? options.rowHeights : Array.isArray(options.rowsMeta) ? options.rowsMeta : [];
  const match = heights.find((item, index) => {
    if (typeof item === "number") return index === rowIndex;
    if (!item || typeof item !== "object") return false;
    const start = Math.max(1, Number(item.row ?? item.index ?? item.min ?? item.start ?? rowNumber) || rowNumber);
    const end = Math.max(start, Number(item.max ?? item.end ?? start) || start);
    return rowNumber >= start && rowNumber <= end;
  });
  const height = typeof match === "number" ? match : match?.height ?? match?.size;
  if (!Number.isFinite(Number(height))) return "";
  const hidden = match && typeof match === "object" && match.hidden ? ' hidden="1"' : "";
  return ` ht="${Math.max(0.1, Number(height))}" customHeight="1"${hidden}`;
}

function xlsxSheetViews(options = {}) {
  const freeze = options.freeze || options.freezePanes || {};
  const cell = typeof freeze.cell === "string" ? parseCellAddress(freeze.cell) : null;
  const xSplit = cell ? cell.col : Math.max(0, Number(freeze.columns || options.freezeColumns || 0) || 0);
  const ySplit = cell ? cell.row : Math.max(0, Number(freeze.rows || options.freezeRows || 0) || 0);
  const topLeftCell = xSplit || ySplit ? `${columnName(xSplit)}${ySplit + 1}` : "A1";
  const activePane = xSplit && ySplit ? "bottomRight" : xSplit ? "topRight" : ySplit ? "bottomLeft" : "topLeft";
  const pane = xSplit || ySplit ? `<pane${xSplit ? ` xSplit="${xSplit}"` : ""}${ySplit ? ` ySplit="${ySplit}"` : ""} topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/><selection pane="${activePane}" activeCell="${topLeftCell}" sqref="${topLeftCell}"/>` : "";
  const showGridLines = typeof options.showGridlines === "boolean" ? ` showGridLines="${options.showGridlines ? 1 : 0}"` : "";
  const showRowColHeaders = typeof options.showHeadings === "boolean" ? ` showRowColHeaders="${options.showHeadings ? 1 : 0}"` : "";
  const zoomScale = Number.isFinite(Number(options.zoomScale || options.zoom)) ? ` zoomScale="${Math.max(10, Math.min(400, Number(options.zoomScale || options.zoom)))}"` : "";
  return `<sheetViews><sheetView workbookViewId="0"${showGridLines}${showRowColHeaders}${zoomScale}>${pane}</sheetView></sheetViews>`;
}

function columnIndexFromValue(value) {
  if (Number.isFinite(Number(value))) return Math.max(1, Math.trunc(Number(value)));
  const text = String(value || "A").toUpperCase().replace(/[^A-Z]/g, "") || "A";
  let index = 0;
  for (const char of text) index = index * 26 + (char.charCodeAt(0) - 64);
  return Math.max(1, index || 1);
}

function xlsxCols(options = {}) {
  const columns = Array.isArray(options.columns) ? options.columns : Array.isArray(options.columnWidths) ? options.columnWidths : [];
  const parts = columns.map((column, index) => {
    if (typeof column === "number") return `<col min="${index + 1}" max="${index + 1}" width="${Math.max(0.1, column)}" customWidth="1"/>`;
    if (!column || typeof column !== "object") return "";
    const min = columnIndexFromValue(column.min ?? column.start ?? column.column ?? column.index ?? index + 1);
    const max = columnIndexFromValue(column.max ?? column.end ?? column.column ?? column.index ?? min);
    const width = Math.max(0.1, Number(column.width || column.size || 12));
    const hidden = column.hidden ? ' hidden="1"' : "";
    return `<col min="${Math.min(min, max)}" max="${Math.max(min, max)}" width="${width}" customWidth="1"${hidden}/>`;
  }).filter(Boolean).join("");
  return parts ? `<cols>${parts}</cols>` : "";
}

function xlsxMergeCells(options = {}) {
  const merges = (Array.isArray(options.merges) ? options.merges : Array.isArray(options.mergeCells) ? options.mergeCells : []).map((merge) => typeof merge === "string" ? merge : merge?.ref || merge?.range || "").filter(Boolean).slice(0, 500);
  return merges.length ? `<mergeCells count="${merges.length}">${merges.map((ref) => `<mergeCell ref="${xml(ref)}"/>`).join("")}</mergeCells>` : "";
}

function xlsxAutoFilter(options = {}) {
  const ref = typeof options.autoFilter === "string" ? options.autoFilter : typeof options.filter === "string" ? options.filter : options.autoFilter?.ref || options.autoFilter?.range || options.filter?.ref || options.filter?.range || "";
  return ref ? `<autoFilter ref="${xml(ref)}"/>` : "";
}

function xlsxPageMargins(options = {}) {
  const margins = options.margins && typeof options.margins === "object" ? options.margins : {};
  const left = asNumber(options.leftMargin ?? margins.left, 0.7);
  const right = asNumber(options.rightMargin ?? margins.right, 0.7);
  const top = asNumber(options.topMargin ?? margins.top, 0.75);
  const bottom = asNumber(options.bottomMargin ?? margins.bottom, 0.75);
  const header = asNumber(options.headerMargin ?? margins.header, 0.3);
  const footer = asNumber(options.footerMargin ?? margins.footer, 0.3);
  return `<pageMargins left="${left}" right="${right}" top="${top}" bottom="${bottom}" header="${header}" footer="${footer}"/>`;
}

function xlsxPrintOptions(options = {}) {
  const attrs = [];
  if (typeof options.showGridlines === "boolean") attrs.push(`gridLines="${options.showGridlines ? 1 : 0}"`);
  if (typeof options.showHeadings === "boolean") attrs.push(`headings="${options.showHeadings ? 1 : 0}"`);
  if (typeof options.centerHorizontally === "boolean") attrs.push(`horizontalCentered="${options.centerHorizontally ? 1 : 0}"`);
  if (typeof options.centerVertically === "boolean") attrs.push(`verticalCentered="${options.centerVertically ? 1 : 0}"`);
  return attrs.length ? `<printOptions ${attrs.join(" ")}/>` : "";
}

function xlsxPageSetup(options = {}) {
  const attrs = [];
  const orientation = String(options.orientation || "").toLowerCase();
  if (orientation === "landscape" || orientation === "portrait") attrs.push(`orientation="${orientation}"`);
  const paperMap = { letter: 1, legal: 5, a4: 9 };
  const paperSize = paperMap[String(options.paperSize || "").toLowerCase()];
  if (paperSize) attrs.push(`paperSize="${paperSize}"`);
  if (Number.isFinite(Number(options.fitToPagesWide))) attrs.push(`fitToWidth="${Math.max(0, Number(options.fitToPagesWide))}"`);
  if (Number.isFinite(Number(options.fitToPagesTall))) attrs.push(`fitToHeight="${Math.max(0, Number(options.fitToPagesTall))}"`);
  if (Number.isFinite(Number(options.scale))) attrs.push(`scale="${Math.max(10, Math.min(400, Number(options.scale)))}"`);
  if (typeof options.blackAndWhite === "boolean") attrs.push(`blackAndWhite="${options.blackAndWhite ? 1 : 0}"`);
  if (typeof options.draftMode === "boolean") attrs.push(`draft="${options.draftMode ? 1 : 0}"`);
  return attrs.length ? `<pageSetup ${attrs.join(" ")}/>` : "";
}

function xlsxLegacyPasswordHash(password = "") {
  let hash = 0;
  const value = String(password || "");
  for (let i = value.length - 1; i >= 0; i -= 1) {
    const rotated = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash = rotated ^ value.charCodeAt(i);
  }
  hash = hash ^ value.length ^ 0xce4b;
  return (hash & 0xffff).toString(16).toUpperCase().padStart(4, "0");
}

function xlsxSheetProtection(options = {}) {
  const protection = options.protection && typeof options.protection === "object" ? options.protection : options;
  const enabled = protection.protect === true || protection.protected === true || options.protect === true;
  if (!enabled) return "";
  const password = protection.password || options.password || "";
  const attrs = [
    'sheet="1"',
    'objects="1"',
    'scenarios="1"',
    `formatCells="${protection.allowFormatCells || options.allowFormatCells ? 0 : 1}"`,
    `sort="${protection.allowSort || options.allowSort ? 0 : 1}"`,
    `autoFilter="${protection.allowAutoFilter || options.allowAutoFilter ? 0 : 1}"`,
  ];
  if (password) attrs.unshift(`password="${xlsxLegacyPasswordHash(password)}"`);
  return `<sheetProtection ${attrs.join(" ")}/>`;
}

function sheetXml(rows = [], options = {}) {
  const normalized = rows.map((row) => Array.isArray(row) ? row : [row]);
  const dimension = normalized.length && normalized[0]?.length ? `A1:${columnName(Math.max(0, ...normalized.map((row) => row.length - 1)))}${normalized.length}` : "A1";
  const headerStyle = options.headerStyle !== false;
  const tableParts = Array.isArray(options.tables) && options.tables.length ? `<tableParts count="${options.tables.length}">${options.tables.map((_table, index) => `<tablePart r:id="rTable${index + 1}"/>`).join("")}</tableParts>` : "";
  const drawing = options.drawingRelId ? `<drawing r:id="${xml(options.drawingRelId)}"/>` : "";
  const legacyDrawing = options.commentsRelId ? `<legacyDrawing r:id="${xml(options.commentsRelId)}"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="${dimension}"/>${xlsxSheetViews(options)}${xlsxCols(options)}<sheetData>${normalized.map((row, rowIndex) => `<row r="${rowIndex + 1}"${xlsxRowHeightAttrs(options, rowIndex)}>${row.map((cell, colIndex) => cellXml(cell, rowIndex, colIndex).replace('<c ', rowIndex === 0 && headerStyle ? '<c s="1" ' : '<c ')).join("")}</row>`).join("")}</sheetData>${xlsxSheetProtection(options)}${xlsxMergeCells(options)}${xlsxAutoFilter(options)}${xlsxDataValidations(options.validations || [])}${xlsxConditionalFormats(options.conditionalFormats || [])}${xlsxHyperlinksXml(options.hyperlinkRefs || [])}${xlsxPrintOptions(options)}${xlsxPageMargins(options)}${xlsxPageSetup(options)}${drawing}${legacyDrawing}${tableParts}</worksheet>`;
}

function sanitizeSheetName(name, index) {
  return String(name || `Sheet${index + 1}`).replace(/[\\/?*\[\]:]/g, "-").slice(0, 31) || `Sheet${index + 1}`;
}

function xlsxSheetFormulaName(name = "Sheet1") {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function xlsxDefinedName(value = "NamedRange") {
  const cleaned = String(value || "NamedRange").replace(/[^A-Za-z0-9_.\\]/g, "_").replace(/^[^A-Za-z_\\]+/, "_").slice(0, 255);
  return cleaned || "NamedRange";
}

function xlsxDefinedNameReference(item = {}, sheets = [], fallbackSheetIndex = null) {
  const rawRef = item.ref || item.reference || item.address || item.range || "A1";
  const ref = String(rawRef).includes("!") ? String(rawRef) : (() => {
    const sheetIndex = Number.isInteger(fallbackSheetIndex) ? fallbackSheetIndex : Math.max(0, Number(item.sheetIndex || 0) || 0);
    const sheetName = item.sheetName || sheets[sheetIndex]?.name || "Sheet1";
    return `${xlsxSheetFormulaName(sanitizeSheetName(sheetName, sheetIndex))}!${rawRef}`;
  })();
  return xml(ref);
}

function xlsxDefinedNames(sheets = [], workbookNamedRanges = []) {
  const names = [];
  const workbookNames = Array.isArray(workbookNamedRanges) ? workbookNamedRanges : [];
  for (const item of workbookNames) {
    if (!item || typeof item !== "object") continue;
    names.push(`<definedName name="${xml(xlsxDefinedName(item.name))}">${xlsxDefinedNameReference(item, sheets)}</definedName>`);
  }
  for (const [index, sheet] of sheets.entries()) {
    const sheetName = xlsxSheetFormulaName(sanitizeSheetName(sheet.name, index));
    if (sheet.printArea) names.push(`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${xml(`${sheetName}!${String(sheet.printArea)}`)}</definedName>`);
    const repeatRows = String(sheet.repeatRows || "").replace(/\$/g, "").trim();
    const repeatColumns = String(sheet.repeatColumns || "").replace(/\$/g, "").trim();
    const titleRefs = [];
    if (repeatRows) titleRefs.push(`${sheetName}!$${repeatRows.replace(":", ":$")}`);
    if (repeatColumns) titleRefs.push(`${sheetName}!$${repeatColumns.replace(":", ":$")}`);
    if (titleRefs.length) names.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${xml(titleRefs.join(","))}</definedName>`);
    const sheetNames = Array.isArray(sheet.namedRanges) ? sheet.namedRanges : [];
    for (const item of sheetNames) {
      if (!item || typeof item !== "object") continue;
      const localSheetId = item.scope === "workbook" ? "" : ` localSheetId="${index}"`;
      names.push(`<definedName name="${xml(xlsxDefinedName(item.name))}"${localSheetId}>${xlsxDefinedNameReference(item, sheets, index)}</definedName>`);
    }
  }
  return names.length ? `<definedNames>${names.join("")}</definedNames>` : "";
}

function xlsxStyles() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

function xlsxTableXml(table, tableId) {
  const ref = table.ref || "A1:B2";
  const columns = Array.isArray(table.columns) && table.columns.length ? table.columns : ["Column1", "Column2"];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tableId}" name="${xml(table.name || `Table${tableId}`)}" displayName="${xml(table.name || `Table${tableId}`)}" ref="${xml(ref)}" totalsRowShown="0"><autoFilter ref="${xml(ref)}"/><tableColumns count="${columns.length}">${columns.map((column, index) => `<tableColumn id="${index + 1}" name="${xml(column)}"/>`).join("")}</tableColumns><tableStyleInfo name="${xml(table.style || "TableStyleMedium2")}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
}

function xlsxDataValidations(validations = []) {
  if (!validations.length) return "";
  return `<dataValidations count="${validations.length}">${validations.map((validation) => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${xml(validation.address || validation.sqref || "A1")}"><formula1>"${xml((validation.values || []).join(","))}"</formula1></dataValidation>`).join("")}</dataValidations>`;
}

function xlsxConditionalFormats(formats = []) {
  return formats.map((format, index) => `<conditionalFormatting sqref="${xml(format.address || format.sqref || "A1")}"><cfRule type="cellIs" priority="${index + 1}" operator="${xml(format.operator || "greaterThan")}"><formula>${xml(String(format.value ?? 0))}</formula><dxf><fill><patternFill patternType="solid"><fgColor rgb="${xml(String(format.fillColor || "FFFFF2CC").replace(/^#/, "FF"))}"/></patternFill></fill></dxf></cfRule></conditionalFormatting>`).join("");
}

function xlsxDocumentProperties(payload = {}) {
  const properties = payload.properties && typeof payload.properties === "object" ? payload.properties : {};
  const title = properties.title || payload.title || String(payload.fileName || "").replace(/\.xlsx$/i, "");
  return {
    title,
    subject: properties.subject || payload.subject || "",
    creator: properties.creator || properties.author || payload.author || "CTRL Add-in",
    keywords: Array.isArray(properties.keywords) ? properties.keywords.join(", ") : properties.keywords || "",
    description: properties.description || properties.comments || "",
    category: properties.category || "",
    company: properties.company || "",
    manager: properties.manager || "",
  };
}

function xlsxCoreProperties(payload = {}) {
  const props = xlsxDocumentProperties(payload);
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(props.title)}</dc:title><dc:subject>${xml(props.subject)}</dc:subject><dc:creator>${xml(props.creator)}</dc:creator><cp:keywords>${xml(props.keywords)}</cp:keywords><dc:description>${xml(props.description)}</dc:description><cp:category>${xml(props.category)}</cp:category><cp:lastModifiedBy>${xml(props.creator)}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
}

function xlsxAppProperties(payload = {}) {
  const props = xlsxDocumentProperties(payload);
  const sheets = Array.isArray(payload.sheets) && payload.sheets.length ? payload.sheets : [{ name: "Sheet1" }];
  const sheetNames = sheets.map((sheet, index) => sanitizeSheetName(sheet?.name, index));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>CTRL Add-in</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">${sheetNames.map((name) => `<vt:lpstr>${xml(name)}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts><Company>${xml(props.company)}</Company><Manager>${xml(props.manager)}</Manager><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>`;
}

function parseCellAddress(address = "E2") {
  const match = String(address || "E2").toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return { col: 4, row: 1 };
  let col = 0;
  for (const char of match[1]) col = col * 26 + (char.charCodeAt(0) - 64);
  return { col: Math.max(0, col - 1), row: Math.max(0, Number(match[2]) - 1) };
}

function xlsxChartFrame(chart = {}, relId = "rChart1", index = 0) {
  const from = parseCellAddress(chart.cell || chart.startCell || `E${2 + index * 16}`);
  const widthPx = Math.round(asNumber(chart.width, 520) * 1.333);
  const heightPx = Math.round(asNumber(chart.height, 300) * 1.333);
  const columns = Math.max(3, Math.ceil(widthPx / 64));
  const rows = Math.max(8, Math.ceil(heightPx / 20));
  const name = xml(chart.title || `Chart ${index + 1}`);
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${from.col + columns}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row + rows}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${index + 100}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

function normalizeXlsxComments(comments = []) {
  return (Array.isArray(comments) ? comments : []).filter((comment) => comment && typeof comment === "object" && (comment.address || comment.cell) && comment.text).slice(0, 200).map((comment) => ({
    address: String(comment.address || comment.cell).toUpperCase(),
    text: String(comment.text || ""),
    author: String(comment.author || "CTRL Add-in"),
    visible: comment.visible === true,
  }));
}

function xlsxCommentsXml(comments = []) {
  const authors = [...new Set(comments.map((comment) => comment.author || "CTRL Add-in"))];
  const authorIndex = (author) => Math.max(0, authors.indexOf(author || "CTRL Add-in"));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors.map((author) => `<author>${xml(author)}</author>`).join("")}</authors><commentList>${comments.map((comment) => `<comment ref="${xml(comment.address)}" authorId="${authorIndex(comment.author)}"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>${xml(comment.text)}</t></r></text></comment>`).join("")}</commentList></comments>`;
}

function xlsxCommentsVml(comments = []) {
  const shapes = comments.map((comment, index) => {
    const from = parseCellAddress(comment.address || "A1");
    const row = from.row;
    const col = from.col;
    return `<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:144pt;height:72pt;z-index:${index + 1};visibility:${comment.visible ? "visible" : "hidden"}" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`;
}

function normalizeXlsxHyperlinks(links = []) {
  return (Array.isArray(links) ? links : []).filter((link) => link && typeof link === "object" && (link.address || link.cell) && (link.url || link.target || link.location)).slice(0, 200).map((link) => ({
    address: String(link.address || link.cell).toUpperCase(),
    target: String(link.url || link.target || link.location || ""),
    tooltip: link.tooltip || link.screenTip || link.text || link.label || link.display || "",
    display: link.display || link.text || link.label || "",
  }));
}

function xlsxHyperlinksXml(links = []) {
  if (!links.length) return "";
  return `<hyperlinks>${links.map((link) => `<hyperlink ref="${xml(link.address)}" r:id="${xml(link.relId)}"${link.display ? ` display="${xml(link.display)}"` : ""}${link.tooltip ? ` tooltip="${xml(link.tooltip)}"` : ""}/>`).join("")}</hyperlinks>`;
}

function xlsxDrawingXml(items = {}) {
  const images = Array.isArray(items.images) ? items.images : Array.isArray(items) ? items : [];
  const charts = Array.isArray(items.charts) ? items.charts : [];
  const anchors = images.map(({ image, relId }, index) => {
    const from = parseCellAddress(image.cell || image.startCell || `E${2 + index * 12}`);
    const widthPx = Math.round(asNumber(image.width, 420) * 1.333);
    const heightPx = Math.round(asNumber(image.height, 240) * 1.333);
    const columns = Math.max(2, Math.ceil(widthPx / 64));
    const rows = Math.max(4, Math.ceil(heightPx / 20));
    const name = xml(image.altText || image.name || `Image ${index + 1}`);
    return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${from.col + columns}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row + rows}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 2}" name="${name}" descr="${name}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
  }).join("");
  const chartAnchors = charts.map(({ chart, relId }, index) => xlsxChartFrame(chart, relId, index)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}${chartAnchors}</xdr:wsDr>`;
}

export function createGeneratedXlsx(payload = {}) {
  const sheets = Array.isArray(payload.sheets) && payload.sheets.length ? payload.sheets.slice(0, 30) : [{ name: "Sheet1", rows: payload.rows || [], images: payload.images || [] }];
  const entries = [];
  const tableEntries = [];
  const drawingEntries = [];
  const commentEntries = [];
  const imageExts = new Set();
  let tableId = 1;
  let mediaIndex = 1;
  let workbookChartId = 1;
  for (const [sheetIndex, sheet] of sheets.entries()) {
    const tables = Array.isArray(sheet.tables) ? sheet.tables : [];
    for (const table of tables) {
      const tablePath = `xl/tables/table${tableId}.xml`;
      tableEntries.push({ sheetIndex, table, tableId, tablePath });
      tableId += 1;
    }
    const images = normalizeGeneratedImages(sheet.images);
    const comments = normalizeXlsxComments(sheet.comments || sheet.notes);
    if (comments.length) commentEntries.push({ sheetIndex, commentId: commentEntries.length + 1, comments });
    const sheetCharts = Array.isArray(sheet.charts) ? sheet.charts.slice(0, 10).filter((chart) => chart && typeof chart === "object") : [];
    let imageRefs = [];
    let chartRefs = [];
    if (images.length) {
      for (const image of images.slice(0, 20)) {
        const bytes = imageBytes(image);
        if (!bytes) continue;
        const ext = imageExt(image.type);
        imageExts.add(ext);
        const mediaName = `image${mediaIndex++}.${ext}`;
        entries.push([`xl/media/${mediaName}`, bytes]);
        imageRefs.push({ image: { ...image, name: image.name || mediaName }, mediaName, relId: `rImage${imageRefs.length + 1}` });
      }
    }
    for (const chart of sheetCharts) {
      const chartNumber = workbookChartId++;
      entries.push([`xl/charts/chart${chartNumber}.xml`, chartXml(chart)]);
      chartRefs.push({ chart, chartNumber, relId: `rChart${chartRefs.length + 1}` });
    }
    if (imageRefs.length || chartRefs.length) drawingEntries.push({ sheetIndex, drawingId: drawingEntries.length + 1, imageRefs, chartRefs });
  }
  entries.push(["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>${[...imageExts].map((ext) => `<Default Extension="${ext}" ContentType="${imageContentType(ext)}"/>`).join("")}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_sheet, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}${drawingEntries.map((entry) => `<Override PartName="/xl/drawings/drawing${entry.drawingId}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join("")}${commentEntries.map((entry) => `<Override PartName="/xl/comments/comment${entry.commentId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`).join("")}${Array.from({ length: workbookChartId - 1 }, (_unused, index) => `<Override PartName="/xl/charts/chart${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join("")}${tableEntries.map((entry) => `<Override PartName="/${entry.tablePath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join("")}</Types>`]);
  entries.push(["_rels/.rels", rels([
    { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml" },
    { id: "rCore", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
    { id: "rApp", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", target: "docProps/app.xml" },
  ])]);
  entries.push(["docProps/core.xml", xlsxCoreProperties(payload)]);
  entries.push(["docProps/app.xml", xlsxAppProperties(payload)]);
  entries.push(["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((sheet, index) => `<sheet name="${xml(sanitizeSheetName(sheet.name, index))}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets>${xlsxDefinedNames(sheets, payload.namedRanges)}</workbook>`]);
  entries.push(["xl/_rels/workbook.xml.rels", rels([
    ...sheets.map((_sheet, index) => ({ id: `rId${index + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: `worksheets/sheet${index + 1}.xml` })),
    { id: "rStyles", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml" },
  ])]);
  entries.push(["xl/styles.xml", xlsxStyles()]);
  sheets.forEach((sheet, index) => {
    const sheetTables = tableEntries.filter((entry) => entry.sheetIndex === index);
    const drawingEntry = drawingEntries.find((entry) => entry.sheetIndex === index);
    const commentEntry = commentEntries.find((entry) => entry.sheetIndex === index);
    const hyperlinks = normalizeXlsxHyperlinks(sheet.links || sheet.hyperlinks);
    const sheetRelationships = sheetTables.map((entry, relIndex) => ({ id: `rTable${relIndex + 1}`, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table", target: `../tables/table${entry.tableId}.xml` }));
    if (drawingEntry) sheetRelationships.push({ id: "rDrawing1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", target: `../drawings/drawing${drawingEntry.drawingId}.xml` });
    if (commentEntry) {
      sheetRelationships.push({ id: "rComments1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", target: `../comments/comment${commentEntry.commentId}.xml` });
      sheetRelationships.push({ id: "rVmlDrawing1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", target: `../drawings/vmlDrawing${commentEntry.commentId}.vml` });
    }
    const hyperlinkRefs = hyperlinks.map((link, linkIndex) => ({ ...link, relId: `rHyperlink${linkIndex + 1}` }));
    for (const link of hyperlinkRefs) sheetRelationships.push({ id: link.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: link.target, targetMode: "External" });
    entries.push([`xl/worksheets/sheet${index + 1}.xml`, sheetXml(sheet.rows || [], { ...sheet, tables: sheetTables.map((entry) => entry.table), drawingRelId: drawingEntry ? "rDrawing1" : "", commentsRelId: commentEntry ? "rVmlDrawing1" : "", hyperlinkRefs })]);
    if (sheetRelationships.length) entries.push([`xl/worksheets/_rels/sheet${index + 1}.xml.rels`, rels(sheetRelationships)]);
  });
  for (const entry of commentEntries) {
    entries.push([`xl/comments/comment${entry.commentId}.xml`, xlsxCommentsXml(entry.comments)]);
    entries.push([`xl/drawings/vmlDrawing${entry.commentId}.vml`, xlsxCommentsVml(entry.comments)]);
  }
  for (const entry of drawingEntries) {
    entries.push([`xl/drawings/drawing${entry.drawingId}.xml`, xlsxDrawingXml({ images: entry.imageRefs, charts: entry.chartRefs })]);
    entries.push([`xl/drawings/_rels/drawing${entry.drawingId}.xml.rels`, rels([
      ...entry.imageRefs.map((imageRef) => ({ id: imageRef.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `../media/${imageRef.mediaName}` })),
      ...entry.chartRefs.map((chartRef) => ({ id: chartRef.relId, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: `../charts/chart${chartRef.chartNumber}.xml` })),
    ])]);
  }
  for (const entry of tableEntries) entries.push([entry.tablePath, xlsxTableXml(entry.table, entry.tableId)]);
  return zip(entries);
}

function scopedRoot(identity = { tenant: "development", subject: "development" }) {
  const scope = crypto.createHash("sha256").update(`${identity.tenant}\0${identity.subject}`).digest("hex");
  return path.join(GENERATED_ROOT, scope);
}

async function removeExpiredArtifacts(root) {
  let names = [];
  try { names = await readdir(root); } catch (error) { if (error?.code === "ENOENT") return; throw error; }
  await Promise.all(names.slice(0, 500).map(async (name) => {
    if (!/^[0-9]+-[0-9a-f-]+-.+\.(docx|pptx|xlsx)$/i.test(name)) return;
    const filePath = path.join(root, name);
    try {
      const info = await stat(filePath);
      if (info.isFile() && Date.now() - info.mtimeMs > GENERATED_RETENTION_MS) await rm(filePath, { force: true });
    } catch (error) { if (error?.code !== "ENOENT") throw error; }
  }));
}

export async function handleGeneratedOffice(req, res, url, identity) {
  const root = scopedRoot(identity);
  try {
    await removeExpiredArtifacts(root);
    if (url.pathname.startsWith("/api/generated/files/")) {
      const id = path.basename(url.pathname);
      const filePath = path.join(root, id);
      const info = await stat(filePath);
      if (!info.isFile()) return sendJson(res, 404, { error: { message: "Generated file not found" } });
      if (Date.now() - info.mtimeMs > GENERATED_RETENTION_MS) return sendJson(res, 410, { error: { message: "Generated file has expired." } });
      const buffer = await readFile(filePath);
      if (buffer.length > MAX_PAYLOAD_BYTES) return sendJson(res, 413, { error: { message: "Generated file exceeds the configured artifact limit." } });
      const downloadName = safeFileName(id.slice(id.indexOf("-") + 1));
      res.writeHead(200, { "content-type": officeContentType(id), "content-disposition": `attachment; filename="${downloadName}"`, "cache-control": "private, no-store", "x-content-type-options": "nosniff", "content-length": buffer.length });
      res.end(buffer);
      return;
    }

    const kind = url.pathname === "/api/generated/pptx" ? "pptx" : url.pathname === "/api/generated/docx" ? "docx" : url.pathname === "/api/generated/xlsx" ? "xlsx" : "";
    if (!kind) return false;
    if (req.method !== "POST") return sendJson(res, 405, { error: { message: "Method not allowed" } });
    const body = await readJson(req);
    if (!body) return sendJson(res, 400, { error: { message: "Missing generated Office payload" } });
    const resolvedBody = await resolveGeneratedPayloadAssets(body, kind);
    const fileName = kind === "pptx"
      ? safeFileName(resolvedBody.fileName || resolvedBody.title || "generated-deck.pptx")
      : safeFileNameWithExtension(resolvedBody.fileName || resolvedBody.title || `generated-${kind}`, kind);
    const id = `${Date.now()}-${crypto.randomUUID()}-${fileName}`;
    const buffer = kind === "pptx" ? createGeneratedPptx(resolvedBody) : kind === "docx" ? createGeneratedDocx(resolvedBody) : createGeneratedXlsx(resolvedBody);
    const packageInfo = validateGeneratedOfficePackage(buffer, kind);
    await mkdir(root, { recursive: true });
    const filePath = path.join(root, id);
    await writeFile(filePath, buffer);
    const expiresAt = new Date(Date.now() + GENERATED_RETENTION_MS).toISOString();
    return sendJson(res, 200, { ok: true, artifactType: kind, fileName, id, size: buffer.length, packageEntryCount: packageInfo.entryCount, expiresAt, retentionMs: GENERATED_RETENTION_MS, downloadUrl: `/api/generated/files/${encodeURIComponent(id)}` });
  } catch (error) {
    return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}


