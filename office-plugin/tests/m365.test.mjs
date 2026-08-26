import { describe, expect, it } from "vitest";
import zlib from "node:zlib";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createAppServer } from "../server/server.mjs";
import { extractOfficeText, rawFileFallback } from "../server/m365.mjs";
import { OFFICE_TOOL_DEFINITIONS } from "../src/lib/tools";

function crc32(buffer) {
  let crc = -1;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ -1) >>> 0;
}

function zipStore(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, contentText] of Object.entries(entries)) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.from(contentText);
    const compressed = zlib.deflateRawSync(content);
    const crc = crc32(content);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    localParts.push(local, nameBuffer, compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(content.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuffer);
    offset += local.length + nameBuffer.length + compressed.length;
  }
  const centralOffset = offset;
  const central = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(entries).length, 8);
  end.writeUInt16LE(Object.keys(entries).length, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(centralOffset, 16);
  return Buffer.concat([...localParts, central, end]);
}

describe("Microsoft 365 delegated context tooling", () => {
  it("exposes M365 auth/search/read tools to the model", () => {
    const names = OFFICE_TOOL_DEFINITIONS.map((tool) => tool.function.name);
    expect(names).toContain("m365_try_office_sso");
    expect(names).toContain("m365_auth_status");
    expect(names).toContain("m365_search_files");
    expect(names).toContain("m365_read_file");
  });

  it("extracts readable text from a downloaded Word document package", () => {
    const docx = zipStore({
      "word/document.xml": '<w:document><w:body><w:p><w:r><w:t>Revenue assumptions</w:t></w:r></w:p><w:p><w:r><w:t>Use 12 percent growth.</w:t></w:r></w:p></w:body></w:document>',
    });
    const text = extractOfficeText(docx, "planning.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(text).toContain("Revenue assumptions");
    expect(text).toContain("Use 12 percent growth");
  });

  it("returns bounded raw context for unsupported text-like and binary files", () => {
    const textFallback = rawFileFallback(Buffer.from("custom format\nrevenue=123"), "notes.custom", "application/x-custom", 2000);
    expect(textFallback).toContain("Raw text fallback");
    expect(textFallback).toContain("revenue=123");

    const binaryFallback = rawFileFallback(Buffer.from([0, 255, 1, 254, 2, 253, 3, 252]), "sample.bin", "application/octet-stream", 2000);
    expect(binaryFallback).toContain("Binary/base64 fallback");
    expect(binaryFallback).toContain("Base64 sample");
  });

  it("returns placeholder configuration guidance when no client id or token source is configured", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "ctrl-m365-config-test-"));
    const previousClient = process.env.MSAL_CLIENT_ID;
    const previousM365Client = process.env.M365_CLIENT_ID;
    const previousToken = process.env.GRAPH_ACCESS_TOKEN;
    const previousCompat = process.env.M365_COMPAT_TOKEN_PATH;
    const previousCache = process.env.M365_TOKEN_CACHE_PATH;
    delete process.env.MSAL_CLIENT_ID;
    delete process.env.M365_CLIENT_ID;
    delete process.env.GRAPH_ACCESS_TOKEN;
    process.env.M365_COMPAT_TOKEN_PATH = path.join(tempDir, "missing-token.json");
    process.env.M365_TOKEN_CACHE_PATH = path.join(tempDir, "missing-cache.json");

    const server = createAppServer();
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const response = await fetch(`http://127.0.0.1:${port}/api/m365/status`);
      const json = await response.json();
      expect(json.authenticated).toBe(false);
      expect(json.authConfig.source).toBe("not_configured");
      expect(json.authRequired.type).toBe("configuration_required");
      expect(json.authRequired.placeholders.MSAL_CLIENT_ID).toBe("00000000-0000-0000-0000-000000000000");
    } finally {
      await new Promise((resolve) => server.close(resolve));
      if (previousClient === undefined) delete process.env.MSAL_CLIENT_ID; else process.env.MSAL_CLIENT_ID = previousClient;
      if (previousM365Client === undefined) delete process.env.M365_CLIENT_ID; else process.env.M365_CLIENT_ID = previousM365Client;
      if (previousToken === undefined) delete process.env.GRAPH_ACCESS_TOKEN; else process.env.GRAPH_ACCESS_TOKEN = previousToken;
      if (previousCompat === undefined) delete process.env.M365_COMPAT_TOKEN_PATH; else process.env.M365_COMPAT_TOKEN_PATH = previousCompat;
      if (previousCache === undefined) delete process.env.M365_TOKEN_CACHE_PATH; else process.env.M365_TOKEN_CACHE_PATH = previousCache;
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("can import an existing compatible m365_token.json cache", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "ctrl-m365-test-"));
    const compatPath = path.join(tempDir, "m365_token.json");
    const cachePath = path.join(tempDir, "addin-token.json");
    const previousCompat = process.env.M365_COMPAT_TOKEN_PATH;
    const previousCache = process.env.M365_TOKEN_CACHE_PATH;
    const previousFetch = globalThis.fetch;
    await writeFile(compatPath, JSON.stringify({ access_token: "existing-token", refresh_token: "refresh", expires_in: 3600, acquired_at: Math.floor(Date.now() / 1000), account: "user@example.com" }), "utf8");
    process.env.M365_COMPAT_TOKEN_PATH = compatPath;
    process.env.M365_TOKEN_CACHE_PATH = cachePath;
    globalThis.fetch = async (input, init) => {
      expect(String(input)).toContain("https://graph.microsoft.com/v1.0/me");
      expect(init?.headers?.authorization).toBe("Bearer existing-token");
      return new Response(JSON.stringify({ displayName: "Test User", userPrincipalName: "user@example.com", id: "1" }), { status: 200, headers: { "content-type": "application/json" } });
    };

    const server = createAppServer();
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const response = await previousFetch(`http://127.0.0.1:${port}/api/m365/status`);
      const json = await response.json();
      expect(json.authenticated).toBe(true);
      expect(json.tokenSource).toBe("compatible_token_cache");
    } finally {
      await new Promise((resolve) => server.close(resolve));
      globalThis.fetch = previousFetch;
      if (previousCompat === undefined) delete process.env.M365_COMPAT_TOKEN_PATH; else process.env.M365_COMPAT_TOKEN_PATH = previousCompat;
      if (previousCache === undefined) delete process.env.M365_TOKEN_CACHE_PATH; else process.env.M365_TOKEN_CACHE_PATH = previousCache;
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("accepts and caches an Office SSO token before falling back to device login", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "ctrl-m365-sso-test-"));
    const previousCache = process.env.M365_SSO_TOKEN_CACHE_PATH;
    const previousFetch = globalThis.fetch;
    process.env.M365_SSO_TOKEN_CACHE_PATH = path.join(tempDir, "sso-cache.json");
    const payload = Buffer.from(JSON.stringify({ aud: "https://graph.microsoft.com", preferred_username: "user@example.com", exp: Math.floor(Date.now() / 1000) + 3600 })).toString("base64url");
    const fakeToken = `header.${payload}.signature`;

    globalThis.fetch = async (input, init) => {
      expect(String(input)).toContain("https://graph.microsoft.com/v1.0/me");
      expect(init?.headers?.authorization).toBe(`Bearer ${fakeToken}`);
      return new Response(JSON.stringify({ displayName: "SSO User", userPrincipalName: "user@example.com", id: "1" }), { status: 200, headers: { "content-type": "application/json" } });
    };

    const server = createAppServer();
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      const ssoResponse = await previousFetch(`http://127.0.0.1:${port}/api/m365/sso`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ accessToken: fakeToken, source: "OfficeRuntime.auth" }),
      });
      const ssoJson = await ssoResponse.json();
      expect(ssoJson.authenticated).toBe(true);
      expect(ssoJson.mode).toBe("direct");

      const statusResponse = await previousFetch(`http://127.0.0.1:${port}/api/m365/status`);
      const statusJson = await statusResponse.json();
      expect(statusJson.authenticated).toBe(true);
      expect(statusJson.tokenSource).toBe("office_sso_direct");
    } finally {
      await new Promise((resolve) => server.close(resolve));
      globalThis.fetch = previousFetch;
      if (previousCache === undefined) delete process.env.M365_SSO_TOKEN_CACHE_PATH; else process.env.M365_SSO_TOKEN_CACHE_PATH = previousCache;
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
