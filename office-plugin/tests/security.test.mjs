import { describe, expect, it } from "vitest";
import { applySecurityHeaders, assertSafeOutboundUrl, authorizeRequest, originAllowed, parseOriginAllowlist, policyAllows, productionConfigErrors, requestOriginAllowed, safeFilename, safeRequestId } from "../server/security.mjs";

describe("security primitives", () => {
  it("normalizes untrusted request IDs before correlation", () => {
    expect(safeRequestId("valid-request_123")).toBe("valid-request_123");
    expect(safeRequestId("bad request with spaces")).not.toBe("bad request with spaces");
    expect(safeRequestId("x".repeat(101))).not.toBe("x".repeat(101));
  });

  it("requires exact origins instead of URL prefixes", () => {
    const allowed = parseOriginAllowlist("https://api.example.com");
    expect(originAllowed(new URL("https://api.example.com/v1"), allowed)).toBe(true);
    expect(originAllowed(new URL("https://api.example.com.attacker.test/v1"), allowed)).toBe(false);
  });

  it("rejects credential-bearing and private outbound URLs", async () => {
    await expect(assertSafeOutboundUrl("https://user:pass@example.com", ["https://example.com"])).rejects.toThrow();
    await expect(assertSafeOutboundUrl("http://127.0.0.1:3000", [], { requireAllowlist: false })).rejects.toThrow();
  });

  it("sanitizes download filenames", () => {
    expect(safeFilename("..\\secret/evil.pptx")).toBe("__secret_evil.pptx");
    expect(safeFilename("", "fallback.txt")).toBe("fallback.txt");
  });

  it("rejects production requests without a bearer token", async () => {
    const previous = process.env.NODE_ENV;
    process.env.NODE_ENV = "production";
    try {
      const result = await authorizeRequest({ headers: {} });
      expect(result.status).toBe(401);
    } finally {
      if (previous === undefined) delete process.env.NODE_ENV; else process.env.NODE_ENV = previous;
    }
  });

  it("supports tenant-level capability policy", () => {
    const previous = process.env.DISABLE_WEB_TENANTS;
    process.env.DISABLE_WEB_TENANTS = "tenant-blocked";
    try {
      expect(policyAllows({ tenant: "tenant-blocked" }, "web")).toBe(false);
      expect(policyAllows({ tenant: "tenant-ok" }, "web")).toBe(true);
    } finally {
      if (previous === undefined) delete process.env.DISABLE_WEB_TENANTS; else process.env.DISABLE_WEB_TENANTS = previous;
    }
  });

  it("does not emit CORS for an untrusted origin", () => {
    const headers = new Map();
    const response = { setHeader(name, value) { headers.set(name, value); } };
    const previous = process.env.ADDIN_ORIGIN;
    process.env.ADDIN_ORIGIN = "https://ctrl.example.com";
    try {
      applySecurityHeaders(response, "request-1234", { requestOrigin: "https://attacker.example" });
      expect(headers.has("access-control-allow-origin")).toBe(false);
    } finally {
      if (previous === undefined) delete process.env.ADDIN_ORIGIN; else process.env.ADDIN_ORIGIN = previous;
    }
  });

  it("rejects hostile browser origins in production but permits non-browser API calls without Origin", () => {
    const previous = process.env.NODE_ENV;
    process.env.NODE_ENV = "production";
    try {
      expect(requestOriginAllowed({ headers: { origin: "https://attacker.example" } }, "https://ctrl.example.com")).toBe(false);
      expect(requestOriginAllowed({ headers: { origin: "https://ctrl.example.com" } }, "https://ctrl.example.com")).toBe(true);
      expect(requestOriginAllowed({ headers: {} }, "https://ctrl.example.com")).toBe(true);
    } finally {
      if (previous === undefined) delete process.env.NODE_ENV; else process.env.NODE_ENV = previous;
    }
  });

  it("requires concrete production identity and origin configuration", () => {
    const previous = { ...process.env };
    process.env.NODE_ENV = "production";
    delete process.env.BYOK_ALLOWED_TARGETS;
    delete process.env.MSAL_CLIENT_ID;
    delete process.env.MSAL_TENANT_ID;
    delete process.env.OFFICE_SSO_RESOURCE;
    delete process.env.ADDIN_ORIGIN;
    try {
      const errors = productionConfigErrors({ verifyToken: async () => ({ subject: "u", tenant: "t" }) });
      expect(errors.length).toBeGreaterThanOrEqual(8);
      // Tenant authorization must fail closed: an unset allowlist previously
      // authorized any tenant that could obtain a valid token.
      expect(errors).toEqual(expect.arrayContaining([expect.stringContaining("AUTHORIZED_TENANTS")]));
    } finally {
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });

  it("requires injected production persistence and tenant policy adapters", async () => {
    const previous = { ...process.env };
    process.env.NODE_ENV = "production";
    process.env.BYOK_ALLOWED_TARGETS = "https://api.example.com";
    process.env.MSAL_CLIENT_ID = "11111111-1111-1111-1111-111111111111";
    process.env.MSAL_TENANT_ID = "22222222-2222-2222-2222-222222222222";
    process.env.OFFICE_SSO_RESOURCE = "api://ctrl.example.com/11111111-1111-1111-1111-111111111111";
    process.env.MSAL_CLIENT_SECRET = "test-secret";
    process.env.ADDIN_ORIGIN = "https://ctrl.example.com";
    delete process.env.SETTINGS_STORE_MODULE;
    delete process.env.TENANT_POLICY_MODULE;
    try {
      const { productionConfigErrors } = await import("../server/security.mjs");
      expect(productionConfigErrors({ verifyToken: async () => ({ subject: "u", tenant: "t" }) })).toEqual(expect.arrayContaining([
        expect.stringContaining("settingsStore"),
        expect.stringContaining("tenant policy"),
      ]));
    } finally {
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });

  it("fails production readiness when legacy Graph token paths or broad scopes are configured", () => {
    const previous = { ...process.env };
    Object.assign(process.env, { NODE_ENV: "production", BYOK_ALLOWED_TARGETS: "https://api.example.com", MSAL_CLIENT_ID: "11111111-1111-1111-1111-111111111111", MSAL_TENANT_ID: "22222222-2222-2222-2222-222222222222", OFFICE_SSO_RESOURCE: "api://ctrl.example.com/11111111-1111-1111-1111-111111111111", MSAL_CLIENT_SECRET: "test-secret", ADDIN_ORIGIN: "https://ctrl.example.com", SETTINGS_STORE_MODULE: "./server/settings-store.production.example.mjs", TENANT_POLICY_MODULE: "./server/tenant-policy.production.example.mjs", RATE_LIMITER_MODULE: "./server/rate-limiter.production.example.mjs", GRAPH_SCOPES: "openid profile User.Read Files.ReadWrite.All", GRAPH_ACCESS_TOKEN: "should-not-be-present" });
    try {
      expect(productionConfigErrors({ verifyToken: async () => ({ subject: "u", tenant: "t" }) })).toEqual(expect.arrayContaining([expect.stringContaining("Legacy Graph token/cache"), expect.stringContaining("Broad Graph scopes")]));
    } finally {
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });
  it("rejects multi-tenant placeholders and self-inconsistent tenant authorization", () => {
    const previous = { ...process.env };
    const base = { NODE_ENV: "production", BYOK_ALLOWED_TARGETS: "https://api.example.com", IMAGE_ALLOWED_TARGETS: "https://images.example.com", WEB_FETCH_ALLOWED_TARGETS: "https://docs.example.com", MSAL_CLIENT_ID: "11111111-1111-1111-1111-111111111111", MSAL_TENANT_ID: "22222222-2222-2222-2222-222222222222", OFFICE_SSO_RESOURCE: "api://ctrl.example.com/11111111-1111-1111-1111-111111111111", MSAL_CLIENT_SECRET: "test-secret", ADDIN_ORIGIN: "https://ctrl.example.com", SETTINGS_STORE_MODULE: "./server/settings-store.production.example.mjs", TENANT_POLICY_MODULE: "./server/tenant-policy.production.example.mjs", RATE_LIMITER_MODULE: "./server/rate-limiter.production.example.mjs" };
    try {
      Object.assign(process.env, base, { AUTHORIZED_TENANTS: "organizations" });
      expect(productionConfigErrors({ verifyToken: async () => ({}) })).toEqual(expect.arrayContaining([expect.stringContaining("concrete tenant ids")]));

      Object.assign(process.env, base, { AUTHORIZED_TENANTS: "33333333-3333-3333-3333-333333333333" });
      expect(productionConfigErrors({ verifyToken: async () => ({}) })).toEqual(expect.arrayContaining([expect.stringContaining("must include MSAL_TENANT_ID")]));

      Object.assign(process.env, base, { AUTHORIZED_TENANTS: "22222222-2222-2222-2222-222222222222" });
      expect(productionConfigErrors({ verifyToken: async () => ({}) }).filter((error) => error.includes("AUTHORIZED_TENANTS"))).toEqual([]);
    } finally {
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });
});