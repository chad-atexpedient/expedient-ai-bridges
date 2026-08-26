import { createServer } from "node:http";
import { mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { createReadStream, existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";
import { handleM365 } from "./m365.mjs";
import { handleFileContext } from "./file-context.mjs";
import { handleImageAsset } from "./image-asset.mjs";
import { handleGeneratedOffice } from "./generated-office.mjs";
import { applySecurityHeaders, assertSafeOutboundUrl, authorizeConfiguredRequest, parseOriginAllowlist, policyAllows, productionConfigErrors, requestOriginAllowed, safeRequestId } from "./security.mjs";
import { createTenantPolicy, policyConfigurationError } from "./policy.mjs";
import { createAuditLogger } from "./observability.mjs";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const defaultDistDir = path.join(root, "dist", "app");
const defaultSettingsPath = path.join(process.env.LOCALAPPDATA || process.env.APPDATA || os.tmpdir(), "ExpedientAIBridges", "office", "shared-settings.json");
const rateBuckets = new Map();

const mime = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".svg", "image/svg+xml"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".webp", "image/webp"],
]);

function sendJson(res, status, body) {
  res.setHeader("content-type", "application/json; charset=utf-8");
  res.writeHead(status);
  res.end(status === 204 ? undefined : JSON.stringify(body));
}

async function readBody(req, limitBytes = Number(process.env.REQUEST_MAX_BYTES || 2_000_000)) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > limitBytes) throw new Error("Request payload is too large.");
    chunks.push(buffer);
  }
  return chunks.length ? Buffer.concat(chunks) : undefined;
}

export function outboundSignal() {
  return AbortSignal.timeout(Number(process.env.OUTBOUND_TIMEOUT_MS || 15000));
}

async function readResponseBufferLimited(response, limitBytes = Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000)) {
  const chunks = [];
  let total = 0;
  for await (const chunk of response.body || []) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > limitBytes) throw new Error("Upstream response is too large.");
    chunks.push(buffer);
  }
  return Buffer.concat(chunks);
}

async function readResponseTextLimited(response, limitBytes) {
  return (await readResponseBufferLimited(response, limitBytes)).toString("utf8");
}

async function readResponseJsonLimited(response, limitBytes) {
  return JSON.parse(await readResponseTextLimited(response, limitBytes));
}
function settingsStorePath(identity = { tenant: "development", subject: "development" }) {
  if (identity.tenant === "development") return path.resolve(process.env.SETTINGS_STORE_PATH || defaultSettingsPath);
  const safe = crypto.createHash("sha256").update(`${identity.tenant}\0${identity.subject}`).digest("hex");
  return path.resolve(process.env.SETTINGS_STORE_DIR || path.dirname(defaultSettingsPath), `${safe}.settings.json`);
}

function rateLimitKey(req, identity, route) {
  const scope = identity?.tenant && identity?.subject ? `${identity.tenant}:${identity.subject}` : `ip:${req.socket.remoteAddress || "unknown"}`;
  return `${scope}:${route}:${Math.floor(Date.now() / 60_000)}`;
}

/**
 * Evict only expired windows.
 *
 * The original implementation deleted every OTHER bucket once the map grew,
 * which meant a single caller could reset every other user's counter simply by
 * flooding the map with fresh keys. Eviction must never be a rate-limit bypass,
 * so current-window buckets are preserved regardless of pressure.
 */
export function pruneRateBuckets(buckets, keepKey, now = Date.now(), maxSize = 5000) {
  if (buckets.size <= maxSize) return buckets;
  const currentWindow = String(Math.floor(now / 60_000));

  // Pass 1: expired windows are always safe to drop.
  for (const candidate of [...buckets.keys()]) {
    if (candidate !== keepKey && !candidate.endsWith(`:${currentWindow}`)) buckets.delete(candidate);
  }
  if (buckets.size <= maxSize) return buckets;

  // Pass 2: still oversized, so the map is being flooded with fresh keys.
  // Evict the LOWEST counters first. Insertion order would shed whichever
  // caller arrived earliest, handing an attacker a way to reset a victim who is
  // near their limit simply by flooding current-window keys. Low counters are
  // the least valuable accounting to lose, and a flooder's keys sit at 1.
  const ranked = [...buckets.entries()]
    .filter(([candidate]) => candidate !== keepKey)
    .sort((left, right) => left[1] - right[1]);
  for (const [candidate] of ranked.slice(0, buckets.size - maxSize)) buckets.delete(candidate);
  return buckets;
}
function checkRateLimit(req, identity, route) {
  const routeLimit = route === "proxy" || route === "web-fetch" || route === "image-asset" ? Number(process.env.RATE_LIMIT_EXPENSIVE_PER_MINUTE || 30) : Number(process.env.RATE_LIMIT_PER_MINUTE || 120);
  const key = rateLimitKey(req, identity, route);
  const next = (rateBuckets.get(key) || 0) + 1;
  rateBuckets.set(key, next);
  pruneRateBuckets(rateBuckets, key);
  return { allowed: next <= routeLimit, limit: routeLimit, remaining: Math.max(0, routeLimit - next) };
}

async function consumeRateLimit(req, identity, route, limiter) {
  if (limiter && typeof limiter.consume === "function") {
    const result = await limiter.consume({ tenant: identity?.tenant, subject: identity?.subject, route, ip: req.socket.remoteAddress || "unknown" });
    return { allowed: result?.allowed !== false, limit: Number(result?.limit || process.env.RATE_LIMIT_PER_MINUTE || 120), remaining: Math.max(0, Number(result?.remaining ?? 0)) };
  }
  return checkRateLimit(req, identity, route);
}

async function readJsonBody(req, limitBytes = 2_000_000) {
  const body = await readBody(req);
  if (!body?.length) return null;
  if (body.length > limitBytes) throw new Error("Settings payload is too large.");
  return JSON.parse(body.toString("utf8"));
}

async function handleSettings(req, res, identity) {
  const storePath = settingsStorePath(identity);

  const safeSettingsResponse = (value) => {
    if (!value || typeof value !== "object") return value;
    const settings = value.settings && typeof value.settings === "object" ? value.settings : value;
    if (process.env.NODE_ENV !== "production") return value;
    return {
      ...value,
      ...(value.settings ? { settings: { ...settings, provider: { ...settings.provider, apiKey: "" } } } : { provider: { ...settings.provider, apiKey: "" } }),
    };
  };

  if (req.method === "GET") {
    try {
      const raw = await readFile(storePath, "utf8");
      return sendJson(res, 200, safeSettingsResponse(JSON.parse(raw)));
    } catch (error) {
      if (error?.code === "ENOENT") return sendJson(res, 204, null);
      return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }

  if (req.method === "PUT") {
    try {
      const json = await readJsonBody(req);
      const settings = json?.settings ?? json;
      if (!settings || typeof settings !== "object") return sendJson(res, 400, { error: { message: "Missing settings object" } });
      if (process.env.NODE_ENV === "production" && settings.provider?.apiKey) {
        return sendJson(res, 400, { error: { message: "Provider API keys cannot be persisted by the production settings service. Use managed credentials or personal local BYOK mode." } });
      }
      await mkdir(path.dirname(storePath), { recursive: true });
      await writeFile(storePath, JSON.stringify({ settings, updatedAt: new Date().toISOString() }, null, 2), "utf8");
      return sendJson(res, 200, { settings: safeSettingsResponse({ settings }).settings });
    } catch (error) {
      return sendJson(res, 400, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }

  if (req.method === "DELETE") {
    try {
      await rm(storePath, { force: true });
      return sendJson(res, 204, null);
    } catch (error) {
      return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }

  res.writeHead(405, { allow: "GET, PUT, DELETE", "content-type": "application/json; charset=utf-8" });
  res.end(JSON.stringify({ error: { message: "Method not allowed" } }));
}

export function forwardedHeaders(req) {
  const headers = new Headers();
  const allowed = new Set(["content-type", "accept", "anthropic-version", "x-provider-authorization", "x-provider-api-key"]);
  const providerHeaderMap = new Map([
    ["x-provider-authorization", "authorization"],
    ["x-provider-api-key", "x-api-key"],
  ]);

  for (const [key, value] of Object.entries(req.headers || {})) {
    const lowerKey = key.toLowerCase();
    if (!value || !allowed.has(lowerKey)) continue;
    const destination = providerHeaderMap.get(lowerKey) || lowerKey;
    headers.set(destination, Array.isArray(value) ? value.join(", ") : value);
  }
  return headers;
}

async function readResponseBodyLimited(response, limitBytes) {
  const chunks = [];
  let total = 0;
  for await (const chunk of response.body || []) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > limitBytes) throw new Error("Upstream response is too large.");
    chunks.push(buffer);
  }
  return Buffer.concat(chunks);
}

async function handleProxy(req, res, url, allowedTargets) {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "access-control-allow-origin": process.env.ADDIN_ORIGIN || "",
      "access-control-allow-methods": "POST, GET, OPTIONS",
      "access-control-allow-headers": "content-type, accept, anthropic-version, x-provider-authorization, x-provider-api-key",
    });
    res.end();
    return;
  }

  if (!["GET", "POST"].includes(req.method)) return sendJson(res, 405, { error: { message: "Proxy method is not allowed." } });

  const target = url.searchParams.get("target");
  if (!target) return sendJson(res, 400, { error: { message: "Missing target" } });

  let targetUrl;
  try {
    targetUrl = await assertSafeOutboundUrl(target, allowedTargets);
  } catch {
    return sendJson(res, 400, { error: { message: "Invalid target URL" } });
  }

  if (!["http:", "https:"].includes(targetUrl.protocol)) {
    return sendJson(res, 400, { error: { message: "Unsupported target protocol" } });
  }

  try {
    const upstream = await fetch(targetUrl, {
      method: req.method,
      headers: forwardedHeaders(req),
      body: await readBody(req, Number(process.env.PROXY_MAX_BYTES || 2_000_000)),
      redirect: "error",
      signal: AbortSignal.timeout(Number(process.env.OUTBOUND_TIMEOUT_MS || 15000)),
    });

    const responseHeaders = { "cache-control": "no-store" };
    upstream.headers.forEach((value, key) => {
      if (!["content-encoding", "transfer-encoding", "connection"].includes(key.toLowerCase())) responseHeaders[key] = value;
    });
    res.writeHead(upstream.status, responseHeaders);
    const data = await readResponseBodyLimited(upstream, Number(process.env.PROXY_MAX_RESPONSE_BYTES || 10_000_000));
    res.end(data);
  } catch (error) {
    sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}


function decodeHtml(value = "") {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDuckDuckGoHtml(html) {
  const results = [];
  const blocks = html.split('result__body').slice(1);
  for (const block of blocks) {
    const linkMatch = block.match(/class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
    if (!linkMatch) continue;
    let url = decodeHtml(linkMatch[1]);
    try {
      const parsed = new URL(url, "https://duckduckgo.com");
      const uddg = parsed.searchParams.get("uddg");
      if (uddg) url = decodeURIComponent(uddg);
    } catch {
      // Keep the original URL when it cannot be normalized.
    }
    const title = decodeHtml(linkMatch[2].replace(/<[^>]+>/g, " "));
    const snippetMatch = block.match(/class="result__snippet"[^>]*>([\s\S]*?)<\/a>|class="result__snippet"[^>]*>([\s\S]*?)<\/div>/i);
    const snippet = decodeHtml((snippetMatch?.[1] || snippetMatch?.[2] || "").replace(/<[^>]+>/g, " "));
    if (url && title) results.push({ title, url, snippet });
    if (results.length >= 8) break;
  }
  return results;
}

async function duckDuckGoHtmlSearch(query) {
  const htmlUrl = new URL("https://html.duckduckgo.com/html/");
  htmlUrl.searchParams.set("q", query);
  const response = await fetch(htmlUrl, { headers: { "user-agent": "Mozilla/5.0", "accept": "text/html" }, signal: outboundSignal() });
  return parseDuckDuckGoHtml(await readResponseTextLimited(response, Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000)));
}
async function handleWebSearch(req, res, url) {
  const query = url.searchParams.get("q")?.trim();
  if (!query) return sendJson(res, 400, { error: { message: "Missing search query" } });

  try {
    const ddgUrl = new URL("https://api.duckduckgo.com/");
    ddgUrl.searchParams.set("q", query);
    ddgUrl.searchParams.set("format", "json");
    ddgUrl.searchParams.set("no_html", "1");
    ddgUrl.searchParams.set("skip_disambig", "1");
    const response = await fetch(ddgUrl, { headers: { "accept": "application/json" }, signal: outboundSignal() });
    const json = await readResponseJsonLimited(response, Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000));
    const related = Array.isArray(json.RelatedTopics) ? json.RelatedTopics.flatMap((topic) => Array.isArray(topic.Topics) ? topic.Topics : [topic]) : [];
    let results = related
      .map((topic) => ({ title: topic.Text || topic.FirstURL || "Result", url: topic.FirstURL || "", snippet: topic.Text || "" }))
      .filter((item) => item.url)
      .slice(0, 8);
    if (!results.length) results = await duckDuckGoHtmlSearch(query);
    sendJson(res, 200, { query, abstract: json.AbstractText || "", abstractUrl: json.AbstractURL || "", results });
  } catch (error) {
    sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}

async function handleSettingsWithStore(req, res, identity, store) {
  const key = { tenant: identity.tenant, subject: identity.subject };
  const safe = (value) => {
    if (!value || typeof value !== "object") return value;
    const settings = value.settings && typeof value.settings === "object" ? value.settings : value;
    return { ...value, ...(value.settings ? { settings: { ...settings, provider: { ...settings.provider, apiKey: "" } } } : { provider: { ...settings.provider, apiKey: "" } }) };
  };
  try {
    if (req.method === "GET") {
      const value = await store.get(key);
      return value == null ? sendJson(res, 204, null) : sendJson(res, 200, safe(value));
    }
    if (req.method === "PUT") {
      const json = await readJsonBody(req);
      const settings = json?.settings ?? json;
      if (!settings || typeof settings !== "object") return sendJson(res, 400, { error: { message: "Missing settings object" } });
      if (settings.provider?.apiKey) return sendJson(res, 400, { error: { message: "Provider API keys cannot be persisted by the managed settings store." } });
      await store.put(key, { settings, updatedAt: new Date().toISOString() });
      return sendJson(res, 200, safe({ settings }));
    }
    if (req.method === "DELETE") { await store.delete(key); return sendJson(res, 204, null); }
    res.writeHead(405, { allow: "GET, PUT, DELETE", "content-type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ error: { message: "Method not allowed" } }));
  } catch (error) {
    return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}

function parseDuckDuckGoVqd(html = "") {
  return html.match(/vqd=['"]([^'"]+)['"]/)?.[1] || html.match(/vqd=([^&\s]+)/)?.[1] || "";
}

export async function duckDuckGoImageSearch(query, maxResults = 8) {
  const searchUrl = new URL("https://duckduckgo.com/");
  searchUrl.searchParams.set("q", query);
  const page = await fetch(searchUrl, { headers: { "user-agent": "Mozilla/5.0", accept: "text/html" }, signal: outboundSignal() });
  const html = await readResponseTextLimited(page, Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000));
  const vqd = parseDuckDuckGoVqd(html);
  if (!vqd) throw new Error("Image search token was not available.");

  const imageUrl = new URL("https://duckduckgo.com/i.js");
  imageUrl.searchParams.set("l", "us-en");
  imageUrl.searchParams.set("o", "json");
  imageUrl.searchParams.set("q", query);
  imageUrl.searchParams.set("vqd", vqd);
  imageUrl.searchParams.set("f", ",,,");
  imageUrl.searchParams.set("p", "1");
  const response = await fetch(imageUrl, {
    headers: {
      "user-agent": "Mozilla/5.0",
      accept: "application/json,text/javascript,*/*;q=0.8",
      referer: searchUrl.href,
    },
    signal: outboundSignal(),
  });
  if (!response.ok) throw new Error(`Image search failed: ${response.status} ${response.statusText}`);
  const json = await readResponseJsonLimited(response, Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000));
  const results = Array.isArray(json.results) ? json.results : [];
  return results.slice(0, maxResults).map((item) => ({
    title: item.title || item.source || query,
    imageUrl: item.image || item.thumbnail || "",
    thumbnailUrl: item.thumbnail || "",
    pageUrl: item.url || item.image || "",
    source: item.source || "",
    width: Number(item.width) || null,
    height: Number(item.height) || null,
  })).filter((item) => item.imageUrl);
}

async function handleWebImageSearch(req, res, url) {
  const query = url.searchParams.get("q")?.trim();
  const count = Math.min(12, Math.max(1, Number(url.searchParams.get("count") || 8)));
  if (!query) return sendJson(res, 400, { error: { message: "Missing image search query" } });
  try {
    const results = await duckDuckGoImageSearch(query, count);
    return sendJson(res, 200, { query, results, note: "Use imageUrl with Office image tools. Review source/copyright suitability before broad production use." });
  } catch (error) {
    return sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}

async function handleWebFetch(req, res, url) {
  const target = url.searchParams.get("url")?.trim();
  if (!target) return sendJson(res, 400, { error: { message: "Missing URL" } });

  let targetUrl;
  try {
    targetUrl = await assertSafeOutboundUrl(target, parseOriginAllowlist(process.env.WEB_FETCH_ALLOWED_TARGETS || ""));
  } catch {
    return sendJson(res, 400, { error: { message: "Invalid URL" } });
  }

  if (!["http:", "https:"].includes(targetUrl.protocol)) return sendJson(res, 400, { error: { message: "Unsupported URL protocol" } });

  try {
    const response = await fetch(targetUrl, { redirect: "error", signal: outboundSignal(), headers: { "accept": "text/html,text/plain,application/json;q=0.8,*/*;q=0.5" } });
    const text = await readResponseTextLimited(response, Number(process.env.WEB_MAX_RESPONSE_BYTES || 2_000_000));
    const cleaned = text
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 20000);
    sendJson(res, 200, { url: targetUrl.href, status: response.status, text: cleaned });
  } catch (error) {
    sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}
async function serveStatic(req, res, url, distDir) {
  const rawPath = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const requested = path.normalize(path.join(distDir, rawPath));
  const relative = path.relative(distDir, requested);
  if (relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) return sendJson(res, 403, { error: { message: "Forbidden" } });

  let filePath = requested;
  if (!existsSync(filePath)) filePath = path.join(distDir, "index.html");

  try {
    const info = await stat(filePath);
    if (!info.isFile()) throw new Error("Not a file");
    res.writeHead(200, { "content-type": mime.get(path.extname(filePath).toLowerCase()) || "application/octet-stream" });
    createReadStream(filePath).pipe(res);
  } catch {
    sendJson(res, 404, { error: { message: "Build output not found. Run npm/pnpm build first." } });
  }
}

async function loadProductionAdapters(options = {}) {
  if (process.env.NODE_ENV !== "production") return { settingsStore: options.settingsStore, policy: options.policy || options.tenantPolicy, rateLimiter: options.rateLimiter };
  let settingsStore = options.settingsStore;
  let rateLimiter = options.rateLimiter;
  let tenantPolicy = options.policy || options.tenantPolicy;
  const modulePath = (name) => process.env[name] ? (process.env[name].startsWith(".") ? new URL(process.env[name], import.meta.url).href : process.env[name]) : null;
  if (process.env.SETTINGS_STORE_MODULE && !settingsStore) {
    const module = await import(modulePath("SETTINGS_STORE_MODULE"));
    const store = module.default || module.settingsStore || module;
    if (!store || typeof store.get !== "function" || typeof store.put !== "function" || typeof store.delete !== "function") throw new Error("SETTINGS_STORE_MODULE must export get, put, and delete.");
    settingsStore = store;
  }
  if (process.env.TENANT_POLICY_MODULE && !tenantPolicy) {
    const module = await import(modulePath("TENANT_POLICY_MODULE"));
    const policy = module.default || module.tenantPolicy || module;
    if (typeof policy !== "function" && typeof policy?.allows !== "function") throw new Error("TENANT_POLICY_MODULE must export a policy function or allows method.");
    tenantPolicy = policy;
  }
  if (process.env.RATE_LIMITER_MODULE && !rateLimiter) {
    const module = await import(modulePath("RATE_LIMITER_MODULE"));
    rateLimiter = module.default || module.rateLimiter || module;
    if (!rateLimiter || typeof rateLimiter.consume !== "function") throw new Error("RATE_LIMITER_MODULE must export a consume method.");
  }
  return { settingsStore, policy: tenantPolicy, rateLimiter };
}

export function createAppServer(options = {}) {
  const distDir = options.distDir || defaultDistDir;
  const allowedTargets = options.allowedTargets || parseOriginAllowlist();
  let policy = createTenantPolicy(options);
  const audit = createAuditLogger(options);
  const configErrors = [...productionConfigErrors(options), policyConfigurationError(options)].filter(Boolean);
  let settingsStore = options.settingsStore;
  let limiter = options.rateLimiter;
  const injectedConfiguration = (options.configurationPromise || loadProductionAdapters(options)).then((adapters) => {
    if (adapters?.settingsStore) settingsStore = adapters.settingsStore;
    if (adapters?.policy) policy = createTenantPolicy({ policy: adapters.policy });
    if (adapters?.rateLimiter) limiter = adapters.rateLimiter;
    return adapters;
  });
  return createServer(async (req, res) => {
    const requestIdValue = safeRequestId(req.headers["x-request-id"]);
    applySecurityHeaders(res, requestIdValue, { requestOrigin: req.headers.origin });
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    if (url.pathname === "/healthz") return sendJson(res, 200, { ok: true });
    if (url.pathname === "/readyz") {
      try { await injectedConfiguration; } catch { return sendJson(res, 503, { ok: false }); }
      return configErrors.length ? sendJson(res, 503, { ok: false }) : sendJson(res, 200, { ok: true });
    }
    if (url.pathname.startsWith("/api/")) {
      if (!requestOriginAllowed(req)) {
        await audit("request.rejected", { requestId: requestIdValue, route: url.pathname, status: 403, reason: "origin" });
        return sendJson(res, 403, { error: { message: "Request origin is not allowed." } });
      }
      const identity = await authorizeConfiguredRequest(req, { verifyToken: options.verifyToken });
      if (!identity.ok) {
        await audit("request.rejected", { requestId: requestIdValue, route: url.pathname, status: identity.status });
        return sendJson(res, identity.status, { error: { message: identity.message } });
      }
      if (configErrors.length) return sendJson(res, 503, { error: { message: "Service configuration is not ready." } });
      try { await injectedConfiguration; } catch { return sendJson(res, 503, { error: { message: "Service configuration is not ready." } }); }
      const route = url.pathname.slice("/api/".length).split("/")[0];
      const rate = await consumeRateLimit(req, identity, route, limiter);
      res.setHeader("x-ratelimit-limit", rate.limit);
      res.setHeader("x-ratelimit-remaining", rate.remaining);
      if (!rate.allowed) { await audit("rate_limited", { requestId: requestIdValue, tenant: identity.tenant, subject: identity.subject, route }); return sendJson(res, 429, { error: { message: "Too many requests. Try again shortly." } }); }
      const allow = async (capability, message) => {
        const allowed = await policy.allows(identity, capability);
        if (!allowed) { await audit("policy.denied", { requestId: requestIdValue, tenant: identity.tenant, route, capability }); return sendJson(res, 403, { error: { message } }); }
        return true;
      };
      if (url.pathname.startsWith("/api/m365/")) return await allow("m365", "Microsoft 365 context is disabled by tenant policy.") ? handleM365(req, res, url, identity) : undefined;
      if (url.pathname === "/api/file-context") return await allow("uploads", "File uploads are disabled by tenant policy.") ? handleFileContext(req, res) : undefined;
      if (url.pathname === "/api/image-asset") return await allow("images", "Image imports are disabled by tenant policy.") ? handleImageAsset(req, res, url) : undefined;
      if (url.pathname.startsWith("/api/generated/")) return await allow("generated", "Generated artifacts are disabled by tenant policy.") ? handleGeneratedOffice(req, res, url, identity) : undefined;
      if (url.pathname === "/api/settings") {
        if (settingsStore) return handleSettingsWithStore(req, res, identity, settingsStore);
        return handleSettings(req, res, identity);
      }
      if (url.pathname === "/api/proxy") return await allow("providers", "Provider access is disabled by tenant policy.") ? handleProxy(req, res, url, allowedTargets) : undefined;
      if (url.pathname === "/api/web-search") return await allow("web", "Web access is disabled by tenant policy.") ? handleWebSearch(req, res, url) : undefined;
      if (url.pathname === "/api/web-image-search") return await allow("web", "Web access is disabled by tenant policy.") ? handleWebImageSearch(req, res, url) : undefined;
      if (url.pathname === "/api/web-fetch") return await allow("web", "Web access is disabled by tenant policy.") ? handleWebFetch(req, res, url) : undefined;
    }
    return serveStatic(req, res, url, distDir);
  });
}

export function startServer(options = {}) {
  const port = Number(options.port || process.env.PORT || 3000);
  const host = options.host || process.env.HOST || "0.0.0.0";
  const server = createAppServer(options);
  server.listen(port, host, () => {
    console.log(`Expedient AI Workspace server listening on http://${host}:${port}`);
    console.log(`Serving ${options.distDir || defaultDistDir}`);
  });
  return server;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  startServer();
}
