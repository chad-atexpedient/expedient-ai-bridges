import { describe, expect, it } from "vitest";
import { createTenantPolicy } from "../server/policy.mjs";
import { safeAuditEvent } from "../server/observability.mjs";

describe("tenant policy and privacy-safe observability", () => {
  it("uses an injected tenant policy rather than global state", async () => {
    const policy = createTenantPolicy({ policy: async ({ identity, capability }) => identity.tenant === "allowed" && capability === "providers" });
    await expect(policy.allows({ tenant: "allowed", subject: "u" }, "providers")).resolves.toBe(true);
    await expect(policy.allows({ tenant: "blocked", subject: "u" }, "providers")).resolves.toBe(false);
    await expect(policy.allows({ tenant: "allowed", subject: "u" }, "web")).resolves.toBe(false);
  });

  it("redacts content and secrets from audit events", () => {
    const record = safeAuditEvent("provider.request", { tenant: "t", subject: "u", route: "/api/proxy", prompt: "secret prompt", apiKey: "sk-secret", content: "document text", durationMs: 42 });
    expect(record).toMatchObject({ event: "provider.request", tenant: "t", subject: "u", route: "/api/proxy", durationMs: 42 });
    expect(JSON.stringify(record)).not.toContain("secret");
    expect(JSON.stringify(record)).not.toContain("document text");
  });
});
