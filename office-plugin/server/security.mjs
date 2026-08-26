import { lookup } from "node:dns/promises";
import { isIP } from "node:net";
import crypto from "node:crypto";

const PRIVATE_RANGES = [
  [/^10\./, "private IPv4"],
  [/^127\./, "loopback"],
  [/^169\.254\./, "link-local"],
  [/^192\.168\./, "private IPv4"],
  [/^172\.(1[6-9]|2\d|3[0-1])\./, "private IPv4"],
  [/^0\./, "unspecified IPv4"],
];

export function safeRequestId(value) {
  const supplied = String(value || "").trim();
  return /^[A-Za-z0-9._-]{8,100}$/.test(supplied) ? supplied : crypto.randomUUID();
}

function requestId(req) {
  const supplied = String(req.headers["x-request-id"] || "").trim();
  return /^[A-Za-z0-9._-]{8,100}$/.test(supplied) ? supplied : crypto.randomUUID();
}

export function isProduction() {
  return process.env.NODE_ENV === "production";
}

export function authRequired(options = {}) {
  return options.requireAuth ?? (isProduction() || process.env.REQUIRE_AUTH === "true");
}

export function requestOriginAllowed(req, expectedOrigin = process.env.ADDIN_ORIGIN || "") {
  const supplied = String(req?.headers?.origin || "").trim();
  if (!isProduction() || !expectedOrigin || !supplied) return true;
  return supplied === expectedOrigin;
}

export async function authorizeRequest(req, options = {}) {
  if (!authRequired(options)) return { ok: true, subject: "development", tenant: "development", requestId: requestId(req) };
  const header = String(req.headers.authorization || "");
  if (!header.startsWith("Bearer ")) return { ok: false, status: 401, message: "Authentication required.", requestId: requestId(req) };
  if (typeof options.verifyToken !== "function") {
    return { ok: false, status: 503, message: "Production token validation is not configured.", requestId: requestId(req) };
  }
  try {
    const identity = await options.verifyToken(header.slice(7));
    if (!identity?.subject || !identity?.tenant) throw new Error("Token identity is incomplete.");
    return { ok: true, ...identity, requestId: requestId(req) };
  } catch {
    return { ok: false, status: 401, message: "Invalid authentication token.", requestId: requestId(req) };
  }
}

let verifierPromise;
async function productionVerifier() {
  if (!verifierPromise) {
    const configured = process.env.AUTH_VALIDATOR_MODULE || "./auth.mjs";
    const moduleUrl = configured.startsWith(".") ? new URL(configured, import.meta.url).href : configured;
    verifierPromise = import(moduleUrl);
  }
  const module = await verifierPromise;
  if (typeof module.verifyToken !== "function") throw new Error("Auth validator must export verifyToken.");
  return module.verifyToken;
}

export async function authorizeConfiguredRequest(req, options = {}) {
  if (!authRequired(options)) return authorizeRequest(req, options);
  try {
    return authorizeRequest(req, { ...options, verifyToken: options.verifyToken || await productionVerifier() });
  } catch {
    return { ok: false, status: 503, message: "Production token validation is not configured." };
  }
}

export function policyAllows(identity, capability) {
  if (!identity || identity.tenant === "development") return true;
  const disabled = (process.env[`DISABLE_${capability.toUpperCase()}_TENANTS`] || "").split(",").map((value) => value.trim()).filter(Boolean);
  return !disabled.includes(identity.tenant);
}

export function applySecurityHeaders(res, requestIdValue, options = {}) {
  const origin = options.origin || process.env.ADDIN_ORIGIN || "";
  const requestOrigin = String(options.requestOrigin || "");
  res.setHeader("x-request-id", requestIdValue);
  res.setHeader("x-content-type-options", "nosniff");
  res.setHeader("referrer-policy", "strict-origin-when-cross-origin");
  res.setHeader("permissions-policy", "camera=(), microphone=(), geolocation=()");
  res.setHeader("content-security-policy", "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; script-src 'self' https://appsforoffice.microsoft.com; connect-src 'self' https://graph.microsoft.com https://login.microsoftonline.com; frame-ancestors https://*.office.com https://*.officeapps.live.com");
  if (isProduction()) res.setHeader("strict-transport-security", "max-age=31536000; includeSubDomains");
  if (origin && requestOrigin === origin) res.setHeader("access-control-allow-origin", origin);
  res.setHeader("vary", "Origin");
}

export function parseOriginAllowlist(raw = process.env.BYOK_ALLOWED_TARGETS || "") {
  return raw.split(",").map((value) => value.trim()).filter(Boolean).map((value) => {
    const url = new URL(value);
    if (!url.hostname || !["http:", "https:"].includes(url.protocol) || url.username || url.password || url.pathname !== "/" || url.search || url.hash) throw new Error(`Invalid outbound origin: ${value}`);
    return url.origin;
  });
}

export function originAllowed(target, allowlist) {
  return allowlist.some((origin) => target.origin === origin);
}

export async function assertSafeOutboundUrl(value, allowlist = [], { requireAllowlist = isProduction() } = {}) {
  const target = new URL(value);
  if (!['http:', 'https:'].includes(target.protocol) || target.username || target.password) throw new Error("Only credential-free HTTP(S) URLs are allowed.");
  if (requireAllowlist && !allowlist.length) throw new Error("Outbound URL allowlist is required in production.");
  if (allowlist.length && !originAllowed(target, allowlist)) throw new Error(`Target origin is not allowed: ${target.origin}`);
  const addresses = await lookup(target.hostname, { all: true, verbatim: true });
  for (const address of addresses) {
    const ip = address.address;
    if (isIP(ip) === 6 && (ip === "::1" || ip.startsWith("fc") || ip.startsWith("fd") || ip.startsWith("fe8") || ip.startsWith("fe9") || ip.startsWith("fea") || ip.startsWith("feb"))) throw new Error("Private or local network targets are not allowed.");
    if (PRIVATE_RANGES.some(([pattern]) => pattern.test(ip))) throw new Error("Private or local network targets are not allowed.");
  }
  return target;
}

export function safeFilename(value, fallback = "download") {
  const name = String(value || fallback).replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/\.\.+/g, "_").trim();
  return (name || fallback).slice(0, 180);
}

export function productionConfigErrors(options = {}) {
  const errors = [];
  const legacyGraphVariables = ["GRAPH_ACCESS_TOKEN", "M365_COMPAT_TOKEN_PATH", "OPENWEBUI_M365_TOKEN_PATH", "M365_TOKEN_CACHE_PATH", "M365_SSO_TOKEN_CACHE_PATH", "M365_DEVICE_FLOW_PATH"];
  const graphScopes = (process.env.GRAPH_SCOPES || "openid profile User.Read Files.Read").split(/[\s,]+/).map((scope) => scope.trim()).filter(Boolean);
  if (isProduction() && !process.env.BYOK_ALLOWED_TARGETS) errors.push("BYOK_ALLOWED_TARGETS is required in production.");
  if (isProduction() && !process.env.IMAGE_ALLOWED_TARGETS) errors.push("IMAGE_ALLOWED_TARGETS is required in production.");
  if (isProduction() && !process.env.WEB_FETCH_ALLOWED_TARGETS) errors.push("WEB_FETCH_ALLOWED_TARGETS is required in production.");
  if (isProduction() && !options.verifyToken && !process.env.AUTH_VALIDATOR_MODULE) errors.push("AUTH_VALIDATOR_MODULE is required in production.");
  if (isProduction() && (!process.env.MSAL_CLIENT_ID || process.env.MSAL_CLIENT_ID === "00000000-0000-0000-0000-000000000000")) errors.push("MSAL_CLIENT_ID is required in production.");
  if (isProduction() && (!process.env.MSAL_TENANT_ID || process.env.MSAL_TENANT_ID === "common")) errors.push("A concrete MSAL_TENANT_ID is required in production.");
  if (isProduction()) {
    const authorizedTenants = (process.env.AUTHORIZED_TENANTS || "").split(",").map((value) => value.trim()).filter(Boolean);
    if (!authorizedTenants.length) errors.push("AUTHORIZED_TENANTS is required in production so tenant authorization fails closed.");
    else if (authorizedTenants.some((tenant) => ["common", "organizations", "consumers", "*"].includes(tenant))) errors.push("AUTHORIZED_TENANTS must name concrete tenant ids, not multi-tenant placeholders.");
    else if (process.env.MSAL_TENANT_ID && !authorizedTenants.includes(process.env.MSAL_TENANT_ID)) errors.push("AUTHORIZED_TENANTS must include MSAL_TENANT_ID or the deployment authorizes no tenant it can sign in.");
  }
  if (isProduction() && (!process.env.OFFICE_SSO_RESOURCE || process.env.OFFICE_SSO_RESOURCE.includes("00000000-0000-0000-0000-000000000000"))) errors.push("OFFICE_SSO_RESOURCE is required in production.");
  if (isProduction() && !process.env.MSAL_CLIENT_SECRET) errors.push("MSAL_CLIENT_SECRET is required for the production Graph OBO exchange.");
  if (isProduction() && legacyGraphVariables.some((name) => process.env[name])) errors.push("Legacy Graph token/cache escape hatches are prohibited in production.");
  if (isProduction() && graphScopes.some((scope) => ["Files.ReadWrite.All", "Sites.Read.All", "Directory.Read.All"].includes(scope))) errors.push("Broad Graph scopes require an explicit production exception and are not allowed by the default policy.");
  if (isProduction() && !graphScopes.includes("openid")) errors.push("GRAPH_SCOPES must include openid in production.");
  if (isProduction() && !graphScopes.includes("profile")) errors.push("GRAPH_SCOPES must include profile in production.");
  if (isProduction() && !process.env.ADDIN_ORIGIN?.startsWith("https://")) errors.push("ADDIN_ORIGIN must be an HTTPS origin in production.");
  if (isProduction() && !options.settingsStore && !process.env.SETTINGS_STORE_MODULE) errors.push("An injected settingsStore or SETTINGS_STORE_MODULE is required in production.");
  if (isProduction() && !options.policy && !options.tenantPolicy && !process.env.TENANT_POLICY_MODULE) errors.push("An injected tenant policy or TENANT_POLICY_MODULE is required in production.");
  if (isProduction() && !options.rateLimiter && !process.env.RATE_LIMITER_MODULE) errors.push("An injected distributed rateLimiter or RATE_LIMITER_MODULE is required in production.");
  return errors;
}
