import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { handleM365 } from "./server/m365.mjs";
import { handleFileContext } from "./server/file-context.mjs";
import { handleImageAsset } from "./server/image-asset.mjs";
import { handleGeneratedOffice } from "./server/generated-office.mjs";
import { assertSafeOutboundUrl, parseOriginAllowlist } from "./server/security.mjs";

const devSettingsPath = path.join(process.env.LOCALAPPDATA || process.env.APPDATA || os.tmpdir(), "CTRL-BYOK-Office-Addin", "shared-settings.dev.json");

async function sendJson(res: any, status: number, body: unknown) {
  res.statusCode = status;
  res.setHeader("content-type", "application/json; charset=utf-8");
  res.end(status === 204 ? undefined : JSON.stringify(body));
}

async function readJsonBody(req: any, limitBytes = 2_000_000) {
  const chunks: Buffer[] = [];
  for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  const body = chunks.length ? Buffer.concat(chunks) : undefined;
  if (!body?.length) return null;
  if (body.length > limitBytes) throw new Error("Settings payload is too large.");
  return JSON.parse(body.toString("utf8"));
}

function forwardedProviderHeaders(req: any) {
  const headers = new Headers();
  const allowed = new Set(["content-type", "accept", "anthropic-version", "x-provider-authorization", "x-provider-api-key"]);
  const providerHeaderMap = new Map([
    ["x-provider-authorization", "authorization"],
    ["x-provider-api-key", "x-api-key"],
  ]);
  for (const [key, value] of Object.entries(req.headers || {})) {
    const lowerKey = key.toLowerCase();
    if (!value || !allowed.has(lowerKey)) continue;
    headers.set(providerHeaderMap.get(lowerKey) || lowerKey, Array.isArray(value) ? value.join(", ") : String(value));
  }
  return headers;
}

async function handleSettings(req: any, res: any) {
  if (req.method === "GET") {
    try {
      return sendJson(res, 200, JSON.parse(await readFile(devSettingsPath, "utf8")));
    } catch (error: any) {
      if (error?.code === "ENOENT") return sendJson(res, 204, null);
      return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }
  if (req.method === "PUT") {
    try {
      const json = await readJsonBody(req);
      const settings = json?.settings ?? json;
      if (!settings || typeof settings !== "object") return sendJson(res, 400, { error: { message: "Missing settings object" } });
      await mkdir(path.dirname(devSettingsPath), { recursive: true });
      await writeFile(devSettingsPath, JSON.stringify({ settings, updatedAt: new Date().toISOString() }, null, 2), "utf8");
      return sendJson(res, 200, { settings });
    } catch (error) {
      return sendJson(res, 400, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }
  if (req.method === "DELETE") {
    try {
      await rm(devSettingsPath, { force: true });
      return sendJson(res, 204, null);
    } catch (error) {
      return sendJson(res, 500, { error: { message: error instanceof Error ? error.message : String(error) } });
    }
  }
  res.statusCode = 405;
  res.setHeader("allow", "GET, PUT, DELETE");
  res.end(JSON.stringify({ error: { message: "Method not allowed" } }));
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

function parseDuckDuckGoHtml(html: string) {
  const results: Array<{ title: string; url: string; snippet: string }> = [];
  const blocks = html.split("result__body").slice(1);
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

async function duckDuckGoHtmlSearch(query: string) {
  const htmlUrl = new URL("https://html.duckduckgo.com/html/");
  htmlUrl.searchParams.set("q", query);
  const response = await fetch(htmlUrl, { headers: { "user-agent": "Mozilla/5.0", accept: "text/html" } });
  return parseDuckDuckGoHtml(await response.text());
}
async function handleWebSearch(req: any, res: any) {
  const host = req.headers.host || "localhost:3000";
  const requestUrl = new URL(req.url || "", `https://${host}`);
  const query = requestUrl.searchParams.get("q")?.trim();
  if (!query) return sendJson(res, 400, { error: { message: "Missing search query" } });

  try {
    const ddgUrl = new URL("https://api.duckduckgo.com/");
    ddgUrl.searchParams.set("q", query);
    ddgUrl.searchParams.set("format", "json");
    ddgUrl.searchParams.set("no_html", "1");
    ddgUrl.searchParams.set("skip_disambig", "1");
    const response = await fetch(ddgUrl, { headers: { accept: "application/json" } });
    const json: any = await response.json();
    const related = Array.isArray(json.RelatedTopics) ? json.RelatedTopics.flatMap((topic: any) => Array.isArray(topic.Topics) ? topic.Topics : [topic]) : [];
    let results = related
      .map((topic: any) => ({ title: topic.Text || topic.FirstURL || "Result", url: topic.FirstURL || "", snippet: topic.Text || "" }))
      .filter((item: any) => item.url)
      .slice(0, 8);
    if (!results.length) results = await duckDuckGoHtmlSearch(query);
    sendJson(res, 200, { query, abstract: json.AbstractText || "", abstractUrl: json.AbstractURL || "", results });
  } catch (error) {
    sendJson(res, 502, { error: { message: error instanceof Error ? error.message : String(error) } });
  }
}

async function handleWebFetch(req: any, res: any) {
  const host = req.headers.host || "localhost:3000";
  const requestUrl = new URL(req.url || "", `https://${host}`);
  const target = requestUrl.searchParams.get("url")?.trim();
  if (!target) return sendJson(res, 400, { error: { message: "Missing URL" } });

  try {
    const targetUrl = new URL(target);
    if (!["http:", "https:"].includes(targetUrl.protocol)) return sendJson(res, 400, { error: { message: "Unsupported URL protocol" } });
    const response = await fetch(targetUrl, { headers: { accept: "text/html,text/plain,application/json;q=0.8,*/*;q=0.5" } });
    const text = await response.text();
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
function byokProxyPlugin(): Plugin {
  return {
    name: "byok-provider-proxy",
    configureServer(server) {
      server.middlewares.use("/api/m365", async (req, res) => {
        const host = req.headers.host || "localhost:3000";
        const requestUrl = new URL(req.url || "", `https://${host}`);
        requestUrl.pathname = `/api/m365${requestUrl.pathname}`;
        await handleM365(req, res, requestUrl, { tenant: "development", subject: "development" });
      });
      server.middlewares.use("/api/file-context", handleFileContext);
      server.middlewares.use("/api/image-asset", async (req, res) => {
        const host = req.headers.host || "localhost:3000";
        const requestUrl = new URL(req.url || "", `https://${host}`);
        requestUrl.pathname = "/api/image-asset";
        await handleImageAsset(req, res, requestUrl);
      });
      server.middlewares.use("/api/generated", async (req, res) => {
        const host = req.headers.host || "localhost:3000";
        const requestUrl = new URL(req.url || "", `https://${host}`);
        requestUrl.pathname = `/api/generated${requestUrl.pathname}`;
        await handleGeneratedOffice(req, res, requestUrl);
      });
      server.middlewares.use("/api/settings", handleSettings);
      server.middlewares.use("/api/web-search", handleWebSearch);
      server.middlewares.use("/api/web-fetch", handleWebFetch);
      server.middlewares.use("/api/proxy", async (req, res) => {
        try {
          if (!["GET", "POST"].includes(req.method || "")) {
            res.statusCode = 405;
            res.setHeader("content-type", "application/json");
            res.end(JSON.stringify({ error: { message: "Proxy method is not allowed." } }));
            return;
          }
          const host = req.headers.host || "localhost:3000";
          const requestUrl = new URL(req.url || "", `https://${host}`);
          const target = requestUrl.searchParams.get("target");
          if (!target) {
            res.statusCode = 400;
            res.end("Missing target");
            return;
          }

          const targetUrl = await assertSafeOutboundUrl(target, parseOriginAllowlist(process.env.BYOK_ALLOWED_TARGETS || ""), { requireAllowlist: false });

          const chunks: Buffer[] = [];
          let total = 0;
          const maxRequestBytes = Number(process.env.PROXY_MAX_BYTES || 2_000_000);
          for await (const chunk of req) {
            const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
            total += buffer.length;
            if (total > maxRequestBytes) throw new Error("Request payload is too large.");
            chunks.push(buffer);
          }

          const upstream = await fetch(targetUrl, {
            method: req.method,
            redirect: "error",
            signal: AbortSignal.timeout(Number(process.env.OUTBOUND_TIMEOUT_MS || 15000)),
            headers: forwardedProviderHeaders(req),
            body: chunks.length ? Buffer.concat(chunks) : undefined,
          });

          res.statusCode = upstream.status;
          upstream.headers.forEach((value, key) => {
            if (!["content-encoding", "transfer-encoding", "connection"].includes(key.toLowerCase())) {
              res.setHeader(key, value);
            }
          });
          const buffer = Buffer.from(await upstream.arrayBuffer());
          if (buffer.length > Number(process.env.PROXY_MAX_RESPONSE_BYTES || 10_000_000)) throw new Error("Upstream response is too large.");
          res.end(buffer);
        } catch (error) {
          res.statusCode = 502;
          res.setHeader("content-type", "application/json");
          res.end(JSON.stringify({ error: { message: error instanceof Error ? error.message : String(error) } }));
        }
      });
    },
  };
}

export default defineConfig(async () => {
  let httpsOptions = {};
  try {
    httpsOptions = await devCerts.getHttpsServerOptions();
  } catch {
    httpsOptions = {};
  }

  return {
    plugins: [react(), byokProxyPlugin()],
    server: {
      host: "localhost",
      port: 3000,
      https: httpsOptions,
    },
    preview: {
      host: "localhost",
      port: 4173,
    },
    build: {
      outDir: "dist/app",
      sourcemap: true,
    },
    test: {
      include: ["tests/**/*.{test,spec}.{ts,tsx,js,mjs}"],
      exclude: ["node_modules/**", "dist/**", ".asset-scratch/**"],
    },
  };
});


