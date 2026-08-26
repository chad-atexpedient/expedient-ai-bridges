const ALLOWED_EVENTS = new Set([
  "request.authenticated", "request.rejected", "policy.denied", "provider.request", "provider.response",
  "m365.request", "upload.accepted", "upload.rejected", "artifact.created", "artifact.downloaded", "rate_limited",
]);

export function safeAuditEvent(event, fields = {}) {
  const safeFields = {};
  for (const [key, value] of Object.entries(fields)) {
    if (["content", "prompt", "text", "base64", "apiKey", "token", "accessToken", "response"].includes(key)) continue;
    if (typeof value === "string") safeFields[key] = value.slice(0, 300);
    else if (["number", "boolean"].includes(typeof value) || value === null) safeFields[key] = value;
  }
  return {
    event: ALLOWED_EVENTS.has(event) ? event : "request.authenticated",
    timestamp: new Date().toISOString(),
    ...safeFields,
  };
}

export function createAuditLogger(options = {}) {
  const sink = options.auditLog || options.auditLogger;
  return async (event, fields = {}) => {
    const record = safeAuditEvent(event, fields);
    if (typeof sink === "function") return sink(record);
    if (sink && typeof sink.write === "function") return sink.write(record);
    if (process.env.AUDIT_LOG_STDOUT === "true") console.log(JSON.stringify(record));
    return undefined;
  };
}

