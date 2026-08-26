import crypto from "node:crypto";

let jwksCache = { expiresAt: 0, keys: [] };

function base64Json(value) {
  return JSON.parse(Buffer.from(value, "base64url").toString("utf8"));
}

function issuer() {
  const tenant = process.env.MSAL_TENANT_ID || process.env.M365_TENANT_ID;
  if (process.env.ENTRA_ISSUER) return process.env.ENTRA_ISSUER.replace(/\/$/, "");
  if (!tenant || tenant === "common") throw new Error("ENTRA_ISSUER or a concrete MSAL_TENANT_ID is required.");
  return `https://login.microsoftonline.com/${tenant}/v2.0`;
}

async function signingKeys() {
  const now = Date.now();
  if (jwksCache.expiresAt > now && jwksCache.keys.length) return jwksCache.keys;
  const response = await fetch(`${issuer()}/discovery/v2.0/keys`, { signal: AbortSignal.timeout(5000), headers: { accept: "application/json" } });
  if (!response.ok) throw new Error(`Entra signing keys unavailable: ${response.status}`);
  const json = await response.json();
  jwksCache = { keys: Array.isArray(json.keys) ? json.keys : [], expiresAt: now + 15 * 60 * 1000 };
  return jwksCache.keys;
}

export async function verifyToken(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) throw new Error("Malformed JWT.");
  const header = base64Json(parts[0]);
  const claims = base64Json(parts[1]);
  const expectedIssuer = issuer();
  const audience = process.env.ENTRA_AUDIENCE || process.env.OFFICE_SSO_RESOURCE || process.env.MSAL_CLIENT_ID;
  // aud may legitimately be an array; compare membership rather than identity.
  const audiences = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  if (!audience || claims.iss !== expectedIssuer || !audiences.includes(audience)) throw new Error("JWT issuer or audience is invalid.");
  // Reject v1.0 tokens: this deployment validates the v2.0 issuer only, and a
  // v1.0 token carries different issuer/audience semantics.
  if (claims.ver !== undefined && String(claims.ver) !== "2.0") throw new Error("Only v2.0 tokens are accepted.");
  if (header.typ !== undefined && !["JWT", "at+jwt"].includes(String(header.typ))) throw new Error("Unexpected token type.");
  const now = Math.floor(Date.now() / 1000);
  const clockSkew = 60;
  if (header.alg !== "RS256" || !header.kid || !claims.tid || !claims.oid || !claims.sub || !Number.isFinite(Number(claims.iat)) || Number(claims.iat) > now + clockSkew || !Number.isFinite(Number(claims.exp)) || Number(claims.exp) <= now - clockSkew || (claims.nbf !== undefined && Number(claims.nbf) > now + clockSkew)) throw new Error("JWT claims are invalid.");
  const key = (await signingKeys()).find((candidate) => candidate.kid === header.kid);
  if (!key) throw new Error("JWT signing key is unknown.");
  const publicKey = crypto.createPublicKey({ key, format: "jwk" });
  const valid = crypto.verify("RSA-SHA256", Buffer.from(`${parts[0]}.${parts[1]}`), publicKey, Buffer.from(parts[2], "base64url"));
  if (!valid) throw new Error("JWT signature is invalid.");
  // Fail closed: an unset allowlist previously defaulted to the token's own
  // tenant, which authorized every tenant that could obtain a valid token.
  const configuredTenants = (process.env.AUTHORIZED_TENANTS || process.env.MSAL_TENANT_ID || process.env.M365_TENANT_ID || "").split(",").map((value) => value.trim()).filter(Boolean);
  if (!configuredTenants.length) throw new Error("AUTHORIZED_TENANTS must list the tenants permitted to use this deployment.");
  if (configuredTenants.includes("common") || configuredTenants.includes("organizations") || configuredTenants.includes("*")) throw new Error("AUTHORIZED_TENANTS must name concrete tenant ids.");
  if (!configuredTenants.includes(claims.tid)) throw new Error("Tenant is not authorized.");
  return { subject: claims.oid, tenant: claims.tid, name: claims.name || claims.preferred_username || claims.oid, claims };
}
