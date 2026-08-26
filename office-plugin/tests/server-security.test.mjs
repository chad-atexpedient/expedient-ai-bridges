import { describe, expect, it } from "vitest";
import { createAppServer, forwardedHeaders } from "../server/server.mjs";

describe("authenticated server boundary", () => {
  it("forwards only explicit provider headers and never the caller bearer token", () => {
    const headers = forwardedHeaders({ headers: {
      authorization: "Bearer office-token",
      "x-provider-authorization": "Bearer provider-key",
      "x-provider-api-key": "provider-key",
      cookie: "session=secret",
      "x-request-id": "request-id",
      "content-type": "application/json",
    } });
    expect(headers.get("authorization")).toBe("Bearer provider-key");
    expect(headers.get("x-api-key")).toBe("provider-key");
    expect(headers.get("content-type")).toBe("application/json");
    expect(headers.get("cookie")).toBeNull();
    expect(headers.get("x-request-id")).toBeNull();
  });

  it("rejects an Office SSO assertion whose tenant or subject differs from the API identity", async () => {
    const previous = { ...process.env };
    Object.assign(process.env, { NODE_ENV: "production", BYOK_ALLOWED_TARGETS: "https://api.example.com", IMAGE_ALLOWED_TARGETS: "https://images.example.com", WEB_FETCH_ALLOWED_TARGETS: "https://docs.example.com", MSAL_CLIENT_ID: "11111111-1111-1111-1111-111111111111", MSAL_TENANT_ID: "22222222-2222-2222-2222-222222222222", AUTHORIZED_TENANTS: "22222222-2222-2222-2222-222222222222", OFFICE_SSO_RESOURCE: "api://ctrl.example.com/11111111-1111-1111-1111-111111111111", ADDIN_ORIGIN: "https://ctrl.example.com", MSAL_CLIENT_SECRET: "test-secret" });
    const header = Buffer.from(JSON.stringify({ alg: "none" })).toString("base64url");
    const claims = Buffer.from(JSON.stringify({ tid: "wrong-tenant", oid: "wrong-user", exp: Math.floor(Date.now() / 1000) + 3600 })).toString("base64url");
    const server = createAppServer({ allowedTargets: ["https://api.example.com"], verifyToken: async () => ({ subject: "user-1", tenant: "tenant-1" }), policy: async () => true, settingsStore: { async get() { return null; }, async put() {}, async delete() {} }, rateLimiter: { async consume() { return { allowed: true, limit: 120, remaining: 119 }; } } });
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const response = await fetch(`http://127.0.0.1:${port}/api/m365/sso`, { method: "POST", headers: { authorization: "Bearer test-token", "content-type": "application/json" }, body: JSON.stringify({ accessToken: `${header}.${claims}.signature` }) });
      expect(response.status).toBe(403);
      expect((await response.json()).error.message).toMatch(/does not match/i);
    } finally {
      await new Promise((resolve) => server.close(resolve));
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });

  it("passes authenticated identity into scoped settings and rejects disabled capabilities", async () => {
    const previous = process.env.NODE_ENV;
    const previousDisabled = process.env.DISABLE_WEB_TENANTS;
    const previousTargets = process.env.BYOK_ALLOWED_TARGETS;
    const previousImageTargets = process.env.IMAGE_ALLOWED_TARGETS;
    const previousWebTargets = process.env.WEB_FETCH_ALLOWED_TARGETS;
    const previousClient = process.env.MSAL_CLIENT_ID;
    const previousTenant = process.env.MSAL_TENANT_ID;
    const previousResource = process.env.OFFICE_SSO_RESOURCE;
    const previousOrigin = process.env.ADDIN_ORIGIN;
    const previousSecret = process.env.MSAL_CLIENT_SECRET;
    process.env.NODE_ENV = "production";
    process.env.DISABLE_WEB_TENANTS = "tenant-1";
    process.env.BYOK_ALLOWED_TARGETS = "https://api.example.com";
    process.env.IMAGE_ALLOWED_TARGETS = "https://images.example.com";
    process.env.WEB_FETCH_ALLOWED_TARGETS = "https://docs.example.com";
    process.env.MSAL_CLIENT_ID = "11111111-1111-1111-1111-111111111111";
    process.env.MSAL_TENANT_ID = "22222222-2222-2222-2222-222222222222";
    process.env.AUTHORIZED_TENANTS = "22222222-2222-2222-2222-222222222222";
    process.env.OFFICE_SSO_RESOURCE = "api://ctrl.example.com/11111111-1111-1111-1111-111111111111";
    process.env.ADDIN_ORIGIN = "https://ctrl.example.com";
    process.env.MSAL_CLIENT_SECRET = "test-secret";
    const store = new Map();
    const server = createAppServer({
      allowedTargets: ["https://api.example.com"],
      verifyToken: async () => ({ subject: "user-1", tenant: "tenant-1" }),
      policy: async ({ capability }) => capability !== "web",
      settingsStore: {
        async get(key) { return store.get(`${key.tenant}:${key.subject}`) ?? null; },
        async put(key, value) { store.set(`${key.tenant}:${key.subject}`, value); },
        async delete(key) { store.delete(`${key.tenant}:${key.subject}`); },
      },
      rateLimiter: { async consume() { return { allowed: true, limit: 120, remaining: 119 }; } },
    });
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const webResponse = await fetch(`http://127.0.0.1:${port}/api/web-search?q=test`, { headers: { authorization: "Bearer test-token" } });
      expect(webResponse.status).toBe(403);
      const settingsResponse = await fetch(`http://127.0.0.1:${port}/api/settings`, { headers: { authorization: "Bearer test-token" } });
      expect([200, 204]).toContain(settingsResponse.status);
    } finally {
      await new Promise((resolve) => server.close(resolve));
      if (previous === undefined) delete process.env.NODE_ENV; else process.env.NODE_ENV = previous;
      if (previousDisabled === undefined) delete process.env.DISABLE_WEB_TENANTS; else process.env.DISABLE_WEB_TENANTS = previousDisabled;
      if (previousTargets === undefined) delete process.env.BYOK_ALLOWED_TARGETS; else process.env.BYOK_ALLOWED_TARGETS = previousTargets;
      if (previousImageTargets === undefined) delete process.env.IMAGE_ALLOWED_TARGETS; else process.env.IMAGE_ALLOWED_TARGETS = previousImageTargets;
      if (previousWebTargets === undefined) delete process.env.WEB_FETCH_ALLOWED_TARGETS; else process.env.WEB_FETCH_ALLOWED_TARGETS = previousWebTargets;
      if (previousClient === undefined) delete process.env.MSAL_CLIENT_ID; else process.env.MSAL_CLIENT_ID = previousClient;
      if (previousTenant === undefined) delete process.env.MSAL_TENANT_ID; else process.env.MSAL_TENANT_ID = previousTenant;
      if (previousResource === undefined) delete process.env.OFFICE_SSO_RESOURCE; else process.env.OFFICE_SSO_RESOURCE = previousResource;
      if (previousOrigin === undefined) delete process.env.ADDIN_ORIGIN; else process.env.ADDIN_ORIGIN = previousOrigin;
      if (previousSecret === undefined) delete process.env.MSAL_CLIENT_SECRET; else process.env.MSAL_CLIENT_SECRET = previousSecret;
    }
  });

  it("rate-limits by authenticated identity and route", async () => {
    const previous = { ...process.env };
    Object.assign(process.env, {
      NODE_ENV: "production",
      BYOK_ALLOWED_TARGETS: "https://api.example.com", IMAGE_ALLOWED_TARGETS: "https://images.example.com", WEB_FETCH_ALLOWED_TARGETS: "https://docs.example.com",
      MSAL_CLIENT_ID: "11111111-1111-1111-1111-111111111111",
      MSAL_TENANT_ID: "22222222-2222-2222-2222-222222222222", AUTHORIZED_TENANTS: "22222222-2222-2222-2222-222222222222",
      OFFICE_SSO_RESOURCE: "api://ctrl.example.com/11111111-1111-1111-1111-111111111111",
      ADDIN_ORIGIN: "https://ctrl.example.com",
      MSAL_CLIENT_SECRET: "test-secret",
      RATE_LIMIT_PER_MINUTE: "1",
    });
    const server = createAppServer({
      allowedTargets: ["https://api.example.com"],
      verifyToken: async () => ({ subject: "rate-user", tenant: "rate-tenant" }),
      policy: async () => true,
      settingsStore: { async get() { return null; }, async put() {}, async delete() {} },
      rateLimiter: { async consume({}) { const current = (this.count || 0) + 1; this.count = current; return { allowed: current <= 1, limit: 1, remaining: Math.max(0, 1 - current) }; } },
    });
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const headers = { authorization: "Bearer test-token" };
      const first = await fetch(`http://127.0.0.1:${port}/api/settings`, { headers });
      const second = await fetch(`http://127.0.0.1:${port}/api/settings`, { headers });
      expect([200, 204]).toContain(first.status);
      expect(second.status).toBe(429);
      expect(second.headers.get("x-ratelimit-limit")).toBe("1");
    } finally {
      await new Promise((resolve) => server.close(resolve));
      for (const key of Object.keys(process.env)) if (!(key in previous)) delete process.env[key];
      for (const [key, value] of Object.entries(previous)) process.env[key] = value;
    }
  });
});
