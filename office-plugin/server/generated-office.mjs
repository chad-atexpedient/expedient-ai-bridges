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