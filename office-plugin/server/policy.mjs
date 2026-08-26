const CAPABILITIES = ["providers", "web", "uploads", "images", "generated", "m365"];

function environmentPolicy(identity, capability) {
  if (!identity || identity.tenant === "development") return true;
  const disabled = (process.env[`DISABLE_${capability.toUpperCase()}_TENANTS`] || "")
    .split(",").map((value) => value.trim()).filter(Boolean);
  return !disabled.includes(identity.tenant);
}

export function createTenantPolicy(options = {}) {
  const configured = options.policy || options.tenantPolicy;
  return {
    async allows(identity, capability) {
      if (!CAPABILITIES.includes(capability)) return false;
      if (typeof configured === "function") return Boolean(await configured({ identity, capability }));
      if (configured && typeof configured.allows === "function") return Boolean(await configured.allows(identity, capability));
      return environmentPolicy(identity, capability);
    },
    capabilities: CAPABILITIES,
  };
}

export function policyConfigurationError(options = {}) {
  if (process.env.NODE_ENV !== "production") return null;
  if (options.policy || options.tenantPolicy || process.env.TENANT_POLICY_MODULE) return null;
  return "TENANT_POLICY_MODULE or an injected tenant policy is required in production.";
}

