import { describe, expect, it, beforeEach, afterEach, vi } from "vitest";
import crypto from "node:crypto";

// Exercises the real Entra verifier against a locally minted RSA key, with the
// JWKS endpoint stubbed. These cover the authorization boundary that had no
// coverage at all: tenant allowlisting, audience shape, token version, and
// signature/claim validation.

const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
const KID = "test-key-1";
const ISSUER = "https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111/v2.0";
const AUDIENCE = "api://ctrl.example.com/11111111-1111-1111-1111-111111111111";
const HOME_TENANT = "11111111-1111-1111-1111-111111111111";
const OTHER_TENANT = "22222222-2222-2222-2222-222222222222";

function base64url(value) {
  return Buffer.from(value).toString("base64url");
}

function signToken({ header = {}, claims = {}, key = privateKey } = {}) {
  const now = Math.floor(Date.now() / 1000);
  const fullHeader = { alg: "RS256", kid: KID, typ: "JWT", ...header };
  const fullClaims = {
    iss: ISSUER,
    aud: AUDIENCE,
    tid: HOME_TENANT,
    oid: "user-object-id",
    sub: "subject-id",
    ver: "2.0",
    iat: now - 30,
    exp: now + 600,
    ...claims,
  };
  const signingInput = `${base64url(JSON.stringify(fullHeader))}.${base64url(JSON.stringify(fullClaims))}`;
  const signature = crypto.sign("RSA-SHA256", Buffer.from(signingInput), key).toString("base64url");
  return `${signingInput}.${signature}`;
}

let originalFetch;
let savedEnv;
let verifyToken;

beforeEach(async () => {
  savedEnv = { ...process.env };
  process.env.ENTRA_ISSUER = ISSUER;
  process.env.ENTRA_AUDIENCE = AUDIENCE;
  process.env.AUTHORIZED_TENANTS = HOME_TENANT;

  const jwk = publicKey.export({ format: "jwk" });
  originalFetch = globalThis.fetch;
  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    json: async () => ({ keys: [{ ...jwk, kid: KID, alg: "RS256", use: "sig" }] }),
  });

  // Fresh module per test so the JWKS cache never leaks between cases.
  vi.resetModules();
  ({ verifyToken } = await import("../server/auth.mjs"));
});

afterEach(() => {
  globalThis.fetch = originalFetch;
  for (const key of Object.keys(process.env)) if (!(key in savedEnv)) delete process.env[key];
  for (const [key, value] of Object.entries(savedEnv)) process.env[key] = value;
});

describe("Entra token verification", () => {
  it("accepts a well-formed token from an authorized tenant", async () => {
    const identity = await verifyToken(signToken());
    expect(identity.tenant).toBe(HOME_TENANT);
    expect(identity.subject).toBe("user-object-id");
  });

  it("rejects a token from an unauthorized tenant", async () => {
    await expect(verifyToken(signToken({ claims: { tid: OTHER_TENANT } }))).rejects.toThrow(/Tenant is not authorized/);
  });

  it("fails closed when no tenant allowlist is configured", async () => {
    // Regression: the allowlist previously defaulted to the tokens own tid,
    // which authorized every tenant that could obtain a valid token.
    delete process.env.AUTHORIZED_TENANTS;
    delete process.env.MSAL_TENANT_ID;
    delete process.env.M365_TENANT_ID;
    await expect(verifyToken(signToken({ claims: { tid: OTHER_TENANT } }))).rejects.toThrow(/AUTHORIZED_TENANTS must list/);
  });

  it("refuses wildcard tenant configuration", async () => {
    process.env.AUTHORIZED_TENANTS = "common";
    await expect(verifyToken(signToken())).rejects.toThrow(/concrete tenant ids/);
  });

  it("rejects a wrong audience", async () => {
    await expect(verifyToken(signToken({ claims: { aud: "api://someone-else" } }))).rejects.toThrow(/issuer or audience/);
  });

  it("accepts an array audience that contains the configured value", async () => {
    const identity = await verifyToken(signToken({ claims: { aud: ["api://other", AUDIENCE] } }));
    expect(identity.tenant).toBe(HOME_TENANT);
  });

  it("rejects a wrong issuer", async () => {
    await expect(verifyToken(signToken({ claims: { iss: "https://evil.example.com/v2.0" } }))).rejects.toThrow(/issuer or audience/);
  });

  it("rejects v1.0 tokens", async () => {
    await expect(verifyToken(signToken({ claims: { ver: "1.0" } }))).rejects.toThrow(/v2.0 tokens/);
  });

  it("rejects an expired token", async () => {
    const now = Math.floor(Date.now() / 1000);
    await expect(verifyToken(signToken({ claims: { exp: now - 3600, iat: now - 7200 } }))).rejects.toThrow(/claims are invalid/);
  });

  it("rejects a token that is not yet valid", async () => {
    const now = Math.floor(Date.now() / 1000);
    await expect(verifyToken(signToken({ claims: { nbf: now + 3600 } }))).rejects.toThrow(/claims are invalid/);
  });

  it("rejects the alg=none downgrade", async () => {
    await expect(verifyToken(signToken({ header: { alg: "none" } }))).rejects.toThrow(/claims are invalid/);
  });

  it("rejects a signature from an unrelated key", async () => {
    const attacker = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
    await expect(verifyToken(signToken({ key: attacker.privateKey }))).rejects.toThrow(/signature is invalid/);
  });

  it("rejects an unknown signing key id", async () => {
    await expect(verifyToken(signToken({ header: { kid: "unknown-kid" } }))).rejects.toThrow(/signing key is unknown/);
  });

  it("rejects tokens missing identity claims", async () => {
    await expect(verifyToken(signToken({ claims: { oid: undefined } }))).rejects.toThrow(/claims are invalid/);
  });

  it("rejects malformed tokens", async () => {
    await expect(verifyToken("not.a.jwt.at.all")).rejects.toThrow();
    await expect(verifyToken("")).rejects.toThrow(/Malformed JWT/);
  });
});