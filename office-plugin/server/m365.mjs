import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import zlib from "node:zlib";
import { isProduction } from "./security.mjs";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const defaultStateDir = path.join(process.env.LOCALAPPDATA || process.env.APPDATA || os.tmpdir(), "CTRL-BYOK-Office-Addin");
const productionTokenStates = new Map();

function productionState(identity = {}) {
  const key = `${identity.tenant || "unknown"}:${identity.subject || "unknown"}`;
  let state = productionTokenStates.get(key);
  if (!state) {
    state = { ssoToken: null };
    productionTokenStates.set(key, state);
  }
  return state;
}

function tokenPath() {
  return path.resolve(process.env.M365_TOKEN_CACHE_PATH || path.join(defaultStateDir, "m365-token-cache.json"));
}

function pendingPath() {
  return path.resolve(process.env.M365_DEVICE_FLOW_PATH || path.join(defaultStateDir, "m365-device-flow.json"));
}

function ssoTokenPath() {
  return path.resolve(process.env.M365_SSO_TOKEN_CACHE_PATH || path.join(defaultStateDir, "m365-sso-token-cache.json"));
}

function clientId() {
  return process.env.MSAL_CLIENT_ID || process.env.M365_CLIENT_ID || "";
}

function tenantId() {
  return process.env.MSAL_TENANT_ID || process.env.M365_TENANT_ID || "common";
}

function authConfig() {
  return {
    clientId: clientId(),
    tenantId: tenantId(),
    source: process.env.MSAL_CLIENT_ID || process.env.M365_CLIENT_ID ? "environment" : "not_configured",
    scopes: graphScopes().split(" "),
  };
}

function graphScopes() {
  return (process.env.GRAPH_SCOPES || "openid profile User.Read Files.Read").split(/[\s,]+/).map((scope) => scope.trim()).filter(Boolean).join(" ");
}

function clientSecret() {
  return process.env.MSAL_CLIENT_SECRET || process.env.M365_CLIENT_SECRET || "";
}

function sendJson(res, status, body) {
  res.writeHead(status, { "content-type": "application/json; charset=utf-8" });
  res.end(status === 204 ? undefined : JSON.stringify(body));
}

async function readJson(req, limitBytes = 1_000_000) {
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  const body = chunks.length ? Buffer.concat(chunks) : undefined;
  if (!body?.length) return null;
  if (body.length > limitBytes) throw new Error("Request payload is too large.");
  return JSON.parse(body.toString("utf8"));
}

async function readJsonFile(filePath) {
  try {
    return JSON.parse(await readFile(filePath, "utf8"));
  } catch {
    return null;
  }
}

function compatibleTokenCachePaths() {
  return [
    process.env.M365_COMPAT_TOKEN_PATH,
    process.env.OPENWEBUI_M365_TOKEN_PATH,
    path.join(process.cwd(), "m365_token.json"),
    path.join(process.cwd(), "backend", "data", "m365_token.json"),
    path.join(defaultStateDir, "m365_token.json"),
  ].filter(Boolean).map((candidate) => path.resolve(candidate));
}

function normalizeCompatibleToken(raw) {
  if (!raw?.access_token) return null;
  const acquiredAt = Number(raw.acquired_at || 0) * 1000;
  const expiresIn = Number(raw.expires_in || 3600) * 1000;
  return {
    access_token: raw.access_token,
    refresh_token: raw.refresh_token,
    token_type: raw.token_type || "Bearer",
    scope: raw.scope || graphScopes(),
    expires_at: raw.expires_at || (acquiredAt ? acquiredAt + expiresIn : Date.now() + expiresIn),
    account: raw.account,
    imported_from: raw.imported_from,
    cached_at: raw.cached_at || new Date().toISOString(),
  };
}

async function importCompatibleTokenCache() {
  const currentPath = tokenPath();
  for (const candidate of compatibleTokenCachePaths()) {
    if (candidate === currentPath) continue;
    const raw = await readJsonFile(candidate);
    const normalized = normalizeCompatibleToken(raw);
    if (!normalized?.access_token) continue;
    normalized.imported_from = candidate;
    await writeJsonFile(currentPath, normalized);
    return normalized;
  }
  return null;
}

async function writeJsonFile(filePath, value) {
  await mkdir(path.dirname(filePath), { recursive: true, mode: 0o700 });
  await writeFile(filePath, JSON.stringify(value, null, 2), { encoding: "utf8", mode: 0o600 });
}

function assertOfficeTokenMatchesIdentity(accessToken, identity = {}) {
  if (!isProduction()) return;
  const claims = decodeJwtPayload(accessToken);
  const tokenTenant = claims?.tid;
  const tokenSubject = claims?.oid || claims?.sub;
  if (!tokenTenant || !tokenSubject || tokenTenant !== identity.tenant || tokenSubject !== identity.subject) {
    const error = new Error("Office SSO token identity does not match the authenticated request.");
    error.code = "IDENTITY_MISMATCH";
    throw error;
  }
}

function tokenStillValid(token) {
  return token?.access_token && Number(token.expires_at || 0) > Date.now() + 120_000;
}

function decodeJwtPayload(token = "") {
  try {
    const payload = token.split(".")[1];
    if (!payload) return null;
    const normalized = payload.replace(/-/g, "+").replace(/_/g, "/");
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    return JSON.parse(Buffer.from(padded, "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function tokenCacheFromAccessToken(accessToken, source = "office_sso") {
  const claims = decodeJwtPayload(accessToken) || {};
  return {
    access_token: accessToken,
    token_type: "Bearer",
    scope: claims.scp || graphScopes(),
    aud: claims.aud,
    account: claims.preferred_username || claims.upn || claims.email,
    expires_at: claims.exp ? Number(claims.exp) * 1000 : Date.now() + 55 * 60 * 1000,
    cached_at: new Date().toISOString(),
    source,
  };
}

function tokenResponseToCache(json) {
  return {
    access_token: json.access_token,
    refresh_token: json.refresh_token,
    token_type: json.token_type,
    scope: json.scope,
    expires_at: Date.now() + Math.max(60, Number(json.expires_in || 3600)) * 1000,
    cached_at: new Date().toISOString(),
  };
}

async function postForm(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded", accept: "application/json" },
    body: new URLSearchParams(body),
  });
  const json = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = json.error_description || json.error?.message || json.error || `${response.status} ${response.statusText}`;
    const error = new Error(message);
    error.status = response.status;
    error.payload = json;
    throw error;
  }
  return json;
}

async function exchangeOnBehalfOf(assertion) {
  if (!clientId() || !clientSecret()) return null;
  try {
    const json = await postForm(`https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`, {
      client_id: clientId(),
      client_secret: clientSecret(),
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      requested_token_use: "on_behalf_of",
      assertion,
      scope: process.env.GRAPH_OBO_SCOPE || `${GRAPH_ROOT}/.default`,
    });
    const token = tokenResponseToCache(json);
    token.source = "office_sso_obo";
    if (!isProduction()) await writeJsonFile(ssoTokenPath(), token);
    return token;
  } catch {
    return null;
  }
}

async function cacheOfficeSsoToken(accessToken, source = "Office SSO", identity = {}) {
  assertOfficeTokenMatchesIdentity(accessToken, identity);
  const obo = await exchangeOnBehalfOf(accessToken);
  if (tokenStillValid(obo)) {
    if (isProduction()) productionState(identity).ssoToken = obo;
    else await writeJsonFile(ssoTokenPath(), obo);
    return { token: obo, mode: "obo" };
  }

  if (isProduction()) throw new Error("Office SSO could not be exchanged for a delegated Graph token.");

  const direct = tokenCacheFromAccessToken(accessToken, "office_sso_direct");
  await writeJsonFile(ssoTokenPath(), direct);
  return { token: direct, mode: "direct" };
}

async function refreshToken(cache) {
  if (!cache?.refresh_token || !clientId()) return null;
  try {
    const json = await postForm(`https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`, {
      client_id: clientId(),
      grant_type: "refresh_token",
      refresh_token: cache.refresh_token,
      scope: graphScopes(),
    });
    const next = tokenResponseToCache(json);
    await writeJsonFile(tokenPath(), next);
    return next;
  } catch {
    return null;
  }
}

async function startDeviceLogin() {
  if (isProduction()) {
    return {
      type: "configuration_required",
      message: "Device login is disabled in production. Use Office SSO with a configured On-Behalf-Of exchange.",
      requiredEnv: ["MSAL_CLIENT_ID", "MSAL_TENANT_ID", "MSAL_CLIENT_SECRET", "OFFICE_SSO_RESOURCE"],
    };
  }
  if (!clientId()) {
    return {
      type: "configuration_required",
      message: "Microsoft 365 delegated access needs an Entra app registration. Set MSAL_CLIENT_ID on the add-in server, then restart it. If you already have a delegated token cache, set M365_COMPAT_TOKEN_PATH instead.",
      requiredEnv: ["MSAL_CLIENT_ID"],
      optionalEnv: ["MSAL_TENANT_ID", "GRAPH_SCOPES", "M365_COMPAT_TOKEN_PATH", "GRAPH_ACCESS_TOKEN"],
      placeholders: {
        MSAL_CLIENT_ID: "00000000-0000-0000-0000-000000000000",
        MSAL_TENANT_ID: "common",
      },
    };
  }

  const existing = await readJsonFile(pendingPath());
  if (existing?.expires_at && existing.expires_at > Date.now()) return existing;

  const json = await postForm(`https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/devicecode`, {
    client_id: clientId(),
    scope: graphScopes(),
  });
  const pending = {
    type: "device_login_required",
    userCode: json.user_code,
    verificationUri: json.verification_uri || "https://microsoft.com/devicelogin",
    verificationUriComplete: json.verification_uri_complete,
    message: json.message,
    deviceCode: json.device_code,
    interval: Number(json.interval || 5),
    expires_at: Date.now() + Number(json.expires_in || 900) * 1000,
    started_at: new Date().toISOString(),
  };
  await writeJsonFile(pendingPath(), pending);
  return pending;
}

async function pollDeviceLogin() {
  const pending = await readJsonFile(pendingPath());
  if (!pending?.deviceCode) return null;
  if (pending.expires_at <= Date.now()) {
    await rm(pendingPath(), { force: true });
    return null;
  }
  try {
    const json = await postForm(`https://login.microsoftonline.com/${tenantId()}/oauth2/v2.0/token`, {
      client_id: clientId(),
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: pending.deviceCode,
    });
    const token = tokenResponseToCache(json);
    await writeJsonFile(tokenPath(), token);
    await rm(pendingPath(), { force: true });
    return token;
  } catch (error) {
    const code = error?.payload?.error;
    if (["authorization_pending", "slow_down"].includes(code)) return null;
    if (["expired_token", "authorization_declined", "bad_verification_code"].includes(code)) await rm(pendingPath(), { force: true });
    return null;
  }
}

async function getTokenOrAuth(identity = {}) {
  if (isProduction()) {
    const state = productionState(identity);
    if (tokenStillValid(state.ssoToken)) return { ok: true, token: state.ssoToken.access_token, source: "office_sso_obo" };
    return { ok: false, auth: await startDeviceLogin() };
  }
  if (process.env.GRAPH_ACCESS_TOKEN) return { ok: true, token: process.env.GRAPH_ACCESS_TOKEN, source: "GRAPH_ACCESS_TOKEN" };

  const cached = await readJsonFile(tokenPath());
  if (tokenStillValid(cached)) return { ok: true, token: cached.access_token, source: "cache" };

  const ssoCached = await readJsonFile(ssoTokenPath());
  if (tokenStillValid(ssoCached)) return { ok: true, token: ssoCached.access_token, source: ssoCached.source || "office_sso" };

  const imported = await importCompatibleTokenCache();
  if (tokenStillValid(imported)) return { ok: true, token: imported.access_token, source: "compatible_token_cache" };

  const refreshed = await refreshToken(cached);
  if (tokenStillValid(refreshed)) return { ok: true, token: refreshed.access_token, source: "refresh_token" };

  const polled = await pollDeviceLogin();
  if (tokenStillValid(polled)) return { ok: true, token: polled.access_token, source: "device_code" };

  const auth = await startDeviceLogin();
  return { ok: false, auth };
}

async function graphFetch(pathname, token, init = {}) {
  const response = await fetch(`${GRAPH_ROOT}${pathname}`, {
    ...init,
    headers: { accept: "application/json", authorization: `Bearer ${token}`, ...(init.headers || {}) },
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `${response.status} ${response.statusText}`);
  }
  return response;
}

function graphFileSummary(item) {
  return {
    id: item.id,
    name: item.name,
    webUrl: item.webUrl,
    size: item.size,
    mimeType: item.file?.mimeType || null,
    lastModifiedDateTime: item.lastModifiedDateTime,
    parentDriveId: item.parentReference?.driveId,
    parentPath: item.parentReference?.path,
  };
}

function escapeGraphSearchQuery(query) {
  return String(query || "").replace(/'/g, "''");
}

async function searchFiles(query, top = 10, identity = {}) {
  const auth = await getTokenOrAuth(identity);
  if (!auth.ok) return { authRequired: auth.auth };
  const escaped = encodeURIComponent(escapeGraphSearchQuery(query));
  const response = await graphFetch(`/me/drive/root/search(q='${escaped}')?$top=${Math.min(25, Math.max(1, Number(top) || 10))}`, auth.token);
  const json = await response.json();
  return { authenticated: true, tokenSource: auth.source, results: (json.value || []).filter((item) => item.file).map(graphFileSummary) };
}

function decodeXmlEntities(value) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_match, code) => String.fromCharCode(Number(code)));
}

function zipEntries(buffer) {
  const entries = new Map();
  const maxEntries = Number(process.env.OFFICE_ARCHIVE_MAX_ENTRIES || 1000);
  const maxUncompressed = Number(process.env.OFFICE_ARCHIVE_MAX_UNCOMPRESSED_BYTES || 50 * 1024 * 1024);
  let totalUncompressed = 0;
  let eocd = -1;
  for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 65558); i -= 1) {
    if (buffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return entries;
  const total = buffer.readUInt16LE(eocd + 10);
  if (total > maxEntries) throw new Error("Office archive contains too many entries.");
  let offset = buffer.readUInt32LE(eocd + 16);
  for (let i = 0; i < total; i += 1) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const uncompressedSize = buffer.readUInt32LE(offset + 24);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.subarray(offset + 46, offset + 46 + fileNameLength).toString("utf8");
    const normalizedName = name.replace(/\\/g, "/");
    if (normalizedName.startsWith("/") || normalizedName.split("/").includes("..")) throw new Error("Office archive contains an unsafe path.");
    totalUncompressed += uncompressedSize;
    if (totalUncompressed > maxUncompressed) throw new Error("Office archive expands beyond the configured limit.");
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    if (dataStart < 0 || dataStart + compressedSize > buffer.length) throw new Error("Office archive contains an invalid entry boundary.");
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);
    let content = Buffer.alloc(0);
    if (method === 0) content = compressed;
    else if (method === 8) content = zlib.inflateRawSync(compressed, { maxOutputLength: Math.max(0, maxUncompressed - (totalUncompressed - uncompressedSize)) });
    else throw new Error("Office archive uses an unsupported compression method.");
    entries.set(name, content);
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function xmlToText(xml) {
  return decodeXmlEntities(xml
    .replace(/<w:tab\b[^>]*\/>/g, "\t")
    .replace(/<a:br\b[^>]*\/>/g, "\n")
    .replace(/<\/(w:p|a:p|w:tr|row)>/g, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim());
}


function xmlAttr(xml = "", attrName = "") {
  const attributes = [...xml.matchAll(/([A-Za-z0-9_:-]+)="([^"]*)"/g)];
  const wanted = new Set([attrName, attrName.includes(":") ? attrName.split(":").pop() : `w:${attrName}`]);
  for (const match of attributes) {
    const name = match[1];
    if (wanted.has(name) || wanted.has(name.split(":").pop())) return decodeXmlEntities(match[2]);
  }
  return "";
}

function listZipEntries(entries, pattern) {
  return [...entries.keys()].filter((entry) => pattern.test(entry)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function summarizeOpenXmlParts(entries, pattern) {
  return listZipEntries(entries, pattern).map((entry) => entry.split("/").pop()).filter(Boolean);
}

function extractRelationshipTargets(xml = "") {
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((match) => ({
    id: xmlAttr(match[0], "Id"),
    type: xmlAttr(match[0], "Type").split("/").pop(),
    target: xmlAttr(match[0], "Target"),
  })).filter((rel) => rel.id || rel.target);
}

function extractPresentationContext(entries) {
  const slideEntries = listZipEntries(entries, /^ppt\/slides\/slide\d+\.xml$/);
  const lines = [];
  const themeEntries = summarizeOpenXmlParts(entries, /^ppt\/theme\/theme\d+\.xml$/);
  const layoutEntries = summarizeOpenXmlParts(entries, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/);
  const masterEntries = summarizeOpenXmlParts(entries, /^ppt\/slideMasters\/slideMaster\d+\.xml$/);
  const mediaEntries = summarizeOpenXmlParts(entries, /^ppt\/media\//);
  if (themeEntries.length) lines.push(`Themes: ${themeEntries.join(", ")}`);
  if (masterEntries.length) lines.push(`Slide masters: ${masterEntries.length}`);
  if (layoutEntries.length) lines.push(`Slide layouts: ${layoutEntries.length}`);
  const layoutDetails = listZipEntries(entries, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/).map((entry, index) => {
    const xml = entries.get(entry)?.toString("utf8") || "";
    const name = xmlAttr(xml.match(/<p:cSld\b[^>]*>/)?.[0] || "", "name") || entry.split("/").pop();
    const type = xmlAttr(xml.match(/<p:sldLayout\b[^>]*>/)?.[0] || "", "type") || "unknown";
    const placeholders = [...xml.matchAll(/<p:ph\b[^>]*>/g)].map((match) => xmlAttr(match[0], "type") || "content").filter(Boolean);
    return `Layout ${index}: ${name} (${type})${placeholders.length ? ` placeholders: ${placeholders.join(", ")}` : ""}`;
  });
  if (layoutDetails.length) lines.push(`Layout details:\n${layoutDetails.join("\n")}`);
  if (mediaEntries.length) lines.push(`Embedded media: ${mediaEntries.slice(0, 30).join(", ")}${mediaEntries.length > 30 ? " ..." : ""}`);

  for (const [index, entry] of slideEntries.entries()) {
    const xml = entries.get(entry)?.toString("utf8") || "";
    const relsEntry = entry.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const rels = entries.get(relsEntry)?.toString("utf8") || "";
    const notesEntry = `ppt/notesSlides/notesSlide${index + 1}.xml`;
    const notesXml = entries.get(notesEntry)?.toString("utf8") || "";
    const shapeCount = (xml.match(/<p:sp\b/g) || []).length;
    const pictureCount = (xml.match(/<p:pic\b/g) || []).length;
    const graphicFrameCount = (xml.match(/<p:graphicFrame\b/g) || []).length;
    const relationships = extractRelationshipTargets(rels).filter((rel) => ["image", "slideLayout", "chart", "hyperlink"].includes(rel.type));
    lines.push([
      `Slide ${index + 1}:`,
      xmlToText(xml) || "[No readable slide text]",
      `Objects: ${shapeCount} text/shape, ${pictureCount} picture, ${graphicFrameCount} graphic/table/chart frames`,
      relationships.length ? `Relationships: ${relationships.map((rel) => `${rel.type}:${rel.target}`).join(", ")}` : "",
      notesXml ? `Speaker notes: ${xmlToText(notesXml) || "[notes present but no readable text]"}` : "",
    ].filter(Boolean).join("\n"));
  }
  return lines.join("\n\n") || null;
}

function extractWordContext(entries) {
  const lines = [];
  const doc = entries.get("word/document.xml")?.toString("utf8");
  if (doc) lines.push(`Document body:\n${xmlToText(doc)}`);
  const styleXml = entries.get("word/styles.xml")?.toString("utf8") || "";
  const styles = [...styleXml.matchAll(/<w:style\b[^>]*(?:\/[>]|>[\s\S]*?<\/w:style>)/g)].map((match) => xmlAttr(match[0], "w:styleId") || xmlAttr(match[0], "styleId")).filter(Boolean).slice(0, 60);
  if (styles.length) lines.push(`Styles: ${styles.join(", ")}`);
  const comments = entries.get("word/comments.xml")?.toString("utf8");
  if (comments) lines.push(`Comments:\n${xmlToText(comments)}`);
  const headerEntries = listZipEntries(entries, /^word\/header\d+\.xml$/);
  const footerEntries = listZipEntries(entries, /^word\/footer\d+\.xml$/);
  if (headerEntries.length) lines.push(`Headers:\n${headerEntries.map((entry) => xmlToText(entries.get(entry).toString("utf8"))).filter(Boolean).join("\n")}`);
  if (footerEntries.length) lines.push(`Footers:\n${footerEntries.map((entry) => xmlToText(entries.get(entry).toString("utf8"))).filter(Boolean).join("\n")}`);
  const mediaEntries = summarizeOpenXmlParts(entries, /^word\/media\//);
  if (mediaEntries.length) lines.push(`Embedded media: ${mediaEntries.slice(0, 30).join(", ")}${mediaEntries.length > 30 ? " ..." : ""}`);
  return lines.filter(Boolean).join("\n\n") || null;
}

function extractWorkbookContext(entries) {
  const lines = [];
  const workbook = entries.get("xl/workbook.xml")?.toString("utf8");
  if (workbook) lines.push(`Workbook structure:\n${xmlToText(workbook)}`);
  const sheetEntries = listZipEntries(entries, /^xl\/worksheets\/sheet\d+\.xml$/);
  for (const [index, entry] of sheetEntries.entries()) {
    const xml = entries.get(entry)?.toString("utf8") || "";
    const rows = (xml.match(/<row\b/g) || []).length;
    const cells = (xml.match(/<c\b/g) || []).length;
    const mergeCount = (xml.match(/<mergeCell\b/g) || []).length;
    const autoFilter = xml.match(/<autoFilter\b[^>]*ref="([^"]+)"/)?.[1];
    const inlineStrings = [...xml.matchAll(/<is>[\s\S]*?<\/is>/g)].map((match) => xmlToText(match[0])).filter(Boolean).slice(0, 80);
    lines.push(`Worksheet ${index + 1}: ${rows} rows, ${cells} cells${mergeCount ? `, ${mergeCount} merged ranges` : ""}${autoFilter ? `, autofilter ${autoFilter}` : ""}${inlineStrings.length ? `\nInline strings: ${inlineStrings.join(", ")}` : ""}`);
    const sheetRels = entries.get(entry.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels")?.toString("utf8") || "";
    const hyperlinkRels = extractRelationshipTargets(sheetRels).filter((rel) => rel.type === "hyperlink");
    if (hyperlinkRels.length) lines.push(`Hyperlinks: ${hyperlinkRels.map((rel) => rel.target).join(", ")}`);
  }
  const tableEntries = listZipEntries(entries, /^xl\/tables\/table\d+\.xml$/);
  for (const entry of tableEntries) {
    const xml = entries.get(entry)?.toString("utf8") || "";
    const name = xmlAttr(xml, "name") || entry.split("/").pop();
    const ref = xmlAttr(xml, "ref");
    lines.push(`Table: ${name}${ref ? ` (${ref})` : ""}`);
  }
  const sharedStrings = entries.get("xl/sharedStrings.xml")?.toString("utf8");
  if (sharedStrings) lines.push(`Shared strings:\n${xmlToText(sharedStrings)}`);
  const commentEntries = listZipEntries(entries, /^xl\/comments\/comment\d+\.xml$/);
  for (const entry of commentEntries) {
    const xml = entries.get(entry)?.toString("utf8") || "";
    const text = xmlToText(xml);
    if (text) lines.push(`Comments (${entry.split("/").pop()}):\n${text}`);
  }
  const chartEntries = summarizeOpenXmlParts(entries, /^xl\/charts\/chart\d+\.xml$/);
  if (chartEntries.length) lines.push(`Charts: ${chartEntries.join(", ")}`);
  const mediaEntries = summarizeOpenXmlParts(entries, /^xl\/media\//);
  if (mediaEntries.length) lines.push(`Embedded media: ${mediaEntries.slice(0, 30).join(", ")}${mediaEntries.length > 30 ? " ..." : ""}`);
  return lines.filter(Boolean).join("\n\n") || null;
}

function themeColorValue(themeXml = "", tag = "") {
  const match = themeXml.match(new RegExp(`<a:${tag}\\b[\\s\\S]*?<a:(?:srgbClr|sysClr)\\b[^>]*>`, "i"));
  if (!match) return "";
  const value = match[0].match(/\bval="([0-9A-Fa-f]{6})"/)?.[1] || match[0].match(/\blast="([0-9A-Fa-f]{6})"/)?.[1] || "";
  return value ? `#${value.toUpperCase()}` : "";
}

function themeFontValue(themeXml = "", group = "") {
  return themeXml.match(new RegExp(`<a:${group}Font>[\\s\\S]*?<a:latin\\b[^>]*\\btypeface="([^"]+)"`, "i"))?.[1] || "";
}

export function extractPowerPointBrandProfile(buffer) {
  const entries = zipEntries(buffer);
  const themeEntry = listZipEntries(entries, /^ppt\/theme\/theme\d+\.xml$/)[0];
  const themeXml = themeEntry ? entries.get(themeEntry)?.toString("utf8") || "" : "";
  const colorTags = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const colors = Object.fromEntries(colorTags.map((tag) => [tag, themeColorValue(themeXml, tag)]).filter(([, value]) => value));
  const layouts = listZipEntries(entries, /^ppt\/slideLayouts\/slideLayout\d+\.xml$/).map((entry, index) => {
    const xml = entries.get(entry)?.toString("utf8") || "";
    return { index, name: xmlAttr(xml.match(/<p:cSld\b[^>]*>/)?.[0] || "", "name") || entry.split("/").pop()?.replace(/\.xml$/i, "") || `layout-${index + 1}`, type: xmlAttr(xml.match(/<p:sldLayout\b[^>]*>/)?.[0] || "", "type") || "unknown", placeholders: [...xml.matchAll(/<p:ph\b[^>]*>/g)].map((match) => xmlAttr(match[0], "type") || "content").slice(0, 20) };
  });
  const media = listZipEntries(entries, /^ppt\/media\//).slice(0, 50).map((entry) => ({ name: entry.split("/").pop() || entry, type: entry.split(".").pop()?.toLowerCase() || "unknown", bytes: entries.get(entry)?.length || 0, candidate: /logo|brand|mark|icon|seal|wordmark/i.test(entry) }));
  return { source: themeEntry ? themeEntry.split("/").pop() : null, colors, fonts: { major: themeFontValue(themeXml, "major"), minor: themeFontValue(themeXml, "minor") }, layouts, media, guidance: [Object.keys(colors).length ? "Use the extracted theme colors for accents and backgrounds." : "No theme color scheme was found; use a conservative accessible palette.", themeFontValue(themeXml, "major") || themeFontValue(themeXml, "minor") ? "Prefer the extracted theme fonts when the target Office client has them installed." : "No theme fonts were found; use the target Office default font.", media.some((item) => item.candidate) ? "Review media candidates before using a logo; confirm source ownership and placement." : "No filename-based logo candidate was identified; do not invent a logo."] };
}

export function extractOfficeText(buffer, name = "", mimeType = "") {
  const lowerName = name.toLowerCase();
  const lowerMime = mimeType.toLowerCase();
  const isZipOffice = /\.(docx|pptx|xlsx)$/i.test(lowerName) || lowerMime.includes("officedocument");
  if (!isZipOffice) return null;
  const entries = zipEntries(buffer);
  if (lowerName.endsWith(".docx") || lowerMime.includes("wordprocessingml")) return extractWordContext(entries);
  if (lowerName.endsWith(".pptx") || lowerMime.includes("presentationml")) return extractPresentationContext(entries);
  if (lowerName.endsWith(".xlsx") || lowerMime.includes("spreadsheetml")) return extractWorkbookContext(entries);
  return null;
}

export function officePackageHasPart(buffer, partName = "") {
  if (!Buffer.isBuffer(buffer) || !partName) return false;
  return zipEntries(buffer).has(String(partName).replace(/^\/+/, ""));
}

function textLikeMime(mimeType = "", name = "") {
  return mimeType.startsWith("text/") || /\.(txt|csv|tsv|json|md|markdown|log|xml|html)$/i.test(name) || ["application/json", "application/xml"].includes(mimeType);
}

export function rawFileFallback(buffer, name = "", mimeType = "", maxChars = 12000) {
  const bytes = buffer.subarray(0, Math.min(buffer.length, 64_000));
  const utf8 = bytes.toString("utf8");
  const replacementRatio = utf8.length ? (utf8.match(/\uFFFD/g) || []).length / utf8.length : 1;
  if (replacementRatio < 0.05 && /[\S]/.test(utf8)) {
    return [
      `Raw text fallback for ${name || "Microsoft 365 file"}.`,
      `MIME type: ${mimeType || "unknown"}`,
      `Bytes sampled: ${bytes.length} of ${buffer.length}`,
      "Content:",
      utf8.slice(0, Math.max(1000, Number(maxChars) || 12000)),
    ].join("\n");
  }

  const base64 = bytes.toString("base64");
  return [
    `Binary/base64 fallback for ${name || "Microsoft 365 file"}.`,
    `MIME type: ${mimeType || "unknown"}`,
    `Bytes sampled: ${bytes.length} of ${buffer.length}`,
    "The model may identify the file format from this sample or explain which converter is needed. Do not assume the full file is represented if it was truncated.",
    "Base64 sample:",
    base64.slice(0, Math.max(1000, Number(maxChars) || 12000)),
  ].join("\n");
}

async function readFileContext(fileId, maxChars = 12000, identity = {}) {
  const auth = await getTokenOrAuth(identity);
  if (!auth.ok) return { authRequired: auth.auth };
  const metadataResponse = await graphFetch(`/me/drive/items/${encodeURIComponent(fileId)}`, auth.token);
  const metadata = await metadataResponse.json();
  const contentResponse = await graphFetch(`/me/drive/items/${encodeURIComponent(fileId)}/content`, auth.token, { headers: { accept: "application/octet-stream" } });
  const buffer = Buffer.from(await contentResponse.arrayBuffer());
  const name = metadata.name || "M365 file";
  const mimeType = metadata.file?.mimeType || contentResponse.headers.get("content-type") || "";
  let text = textLikeMime(mimeType, name) ? buffer.toString("utf8") : extractOfficeText(buffer, name, mimeType);
  if (!text) text = rawFileFallback(buffer, name, mimeType, maxChars);
  return { authenticated: true, tokenSource: auth.source, file: graphFileSummary(metadata), text: text.slice(0, Math.max(1000, Number(maxChars) || 12000)) };
}

async function status(identity = {}) {
  const auth = await getTokenOrAuth(identity);
  if (!auth.ok) return { authenticated: false, authConfig: authConfig(), authRequired: auth.auth };
  const response = await graphFetch("/me?$select=displayName,userPrincipalName,id", auth.token);
  return { authenticated: true, tokenSource: auth.source, authConfig: authConfig(), user: await response.json() };
}

export async function handleM365(req, res, url, identity = {}) {
  try {
    if (url.pathname === "/api/m365/logout") {
      if (isProduction()) {
        productionTokenStates.delete(`${identity.tenant || "unknown"}:${identity.subject || "unknown"}`);
        return sendJson(res, 204, null);
      }
      await rm(tokenPath(), { force: true });
      await rm(pendingPath(), { force: true });
      await rm(ssoTokenPath(), { force: true });
      return sendJson(res, 204, null);
    }
    if (url.pathname === "/api/m365/sso") {
      if (req.method !== "POST") return sendJson(res, 405, { error: { message: "Method not allowed" } });
      const body = await readJson(req);
      const accessToken = body?.accessToken || body?.token;
      if (!accessToken || typeof accessToken !== "string") return sendJson(res, 400, { error: { message: "Missing Office SSO access token" } });
      const cached = await cacheOfficeSsoToken(accessToken, body?.source || "Office SSO", identity);
      return sendJson(res, 200, {
        authenticated: true,
        tokenSource: cached.token.source,
        mode: cached.mode,
        message: cached.mode === "obo"
          ? "Office SSO token exchanged for Microsoft Graph access."
          : "Office SSO token cached. If Graph rejects it, configure MSAL_CLIENT_SECRET for On-Behalf-Of exchange.",
        claims: {
          aud: cached.token.aud,
          account: cached.token.account,
          expiresAt: cached.token.expires_at,
        },
      });
    }
    if (url.pathname === "/api/m365/status") return sendJson(res, 200, await status(identity));
    if (url.pathname === "/api/m365/device-login") return sendJson(res, 200, { authRequired: await startDeviceLogin() });
    if (url.pathname === "/api/m365/poll") {
      const token = await pollDeviceLogin();
      return sendJson(res, 200, tokenStillValid(token) ? await status(identity) : await status(identity));
    }
    if (url.pathname === "/api/m365/search") {
      const query = url.searchParams.get("q")?.trim() || (await readJson(req))?.query || "";
      if (!query) return sendJson(res, 400, { error: { message: "Missing Microsoft 365 search query" } });
      return sendJson(res, 200, await searchFiles(query, url.searchParams.get("top") || 10, identity));
    }
    if (url.pathname === "/api/m365/read") {
      const body = req.method === "POST" ? await readJson(req) : null;
      const fileId = url.searchParams.get("id") || body?.id || body?.fileId;
      if (!fileId) return sendJson(res, 400, { error: { message: "Missing Microsoft 365 file id" } });
      return sendJson(res, 200, await readFileContext(fileId, url.searchParams.get("maxChars") || body?.maxChars, identity));
    }
    return false;
  } catch (error) {
    if (error?.code === "IDENTITY_MISMATCH") return sendJson(res, 403, { error: { message: error.message } });
    return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}
