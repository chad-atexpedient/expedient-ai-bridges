import { describe, expect, it, vi } from "vitest";
import http from "node:http";
import { duckDuckGoImageSearch, pruneRateBuckets } from "../server/server.mjs";
import { handleFileContext } from "../server/file-context.mjs";

// Regression tests for four server defects found by inspection. Each test was
// verified to FAIL against the original implementation before the fix landed.
describe("rate-limit bucket eviction", () => {
  it("never drops another caller current-window bucket under memory pressure", () => {
    const now = 1_800_000_000_000;
    const window = String(Math.floor(now / 60_000));
    const buckets = new Map();
    // A victim who is already over their limit in the current window.
    const victim = `tenant-a:victim:chat:${window}`;
    buckets.set(victim, 119);
    // An attacker floods the map with fresh current-window keys to force eviction.
    for (let index = 0; index < 6000; index += 1) buckets.set(`tenant-b:flood-${index}:chat:${window}`, 1);
    const attackerKey = `tenant-b:flood-5999:chat:${window}`;

    pruneRateBuckets(buckets, attackerKey, now);

    // The victim counter must survive: eviction is not a rate-limit reset.
    expect(buckets.get(victim)).toBe(119);
    expect(buckets.has(attackerKey)).toBe(true);
  });

  it("evicts stale windows and leaves the map bounded", () => {
    const now = 1_800_000_000_000;
    const window = Math.floor(now / 60_000);
    const buckets = new Map();
    for (let index = 0; index < 5200; index += 1) buckets.set(`t:u${index}:chat:${window - 30}`, 4);
    const keep = `t:current:chat:${window}`;
    buckets.set(keep, 2);

    pruneRateBuckets(buckets, keep, now);

    expect(buckets.size).toBeLessThanOrEqual(5000);
    expect(buckets.get(keep)).toBe(2);
  });
});

describe("streaming size limits", () => {
  it("rejects an oversized file-context upload without buffering the whole body", async () => {
    const previous = process.env.FILE_CONTEXT_MAX_UPLOAD_BYTES;
    process.env.FILE_CONTEXT_MAX_UPLOAD_BYTES = "2048";
    const server = http.createServer((req, res) => { handleFileContext(req, res); });
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    try {
      const { port } = server.address();
      // 4 MB of payload against a ~2 KB declared limit.
      const chunk = "a".repeat(64 * 1024);
      let sentBytes = 0;
      const response = await fetch(`http://127.0.0.1:${port}/api/file-context`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        duplex: "half",
        body: new ReadableStream({
          async pull(controller) {
            if (sentBytes >= 4 * 1024 * 1024) { controller.close(); return; }
            sentBytes += chunk.length;
            controller.enqueue(new TextEncoder().encode(chunk));
          },
        }),
      }).catch((error) => ({ status: 413, aborted: true, error }));
      // Either a clean 413 or a torn-down connection is acceptable; silently
      // accepting the whole 4 MB body is not.
      expect([413, 400]).toContain(response.status);
      if (!response.aborted) expect((await response.json()).error.message).toMatch(/too large/i);
      // Proves the server stopped reading instead of draining the full upload.
      expect(sentBytes).toBeLessThan(4 * 1024 * 1024);
    } finally {
      await new Promise((resolve) => server.close(resolve));
      if (previous === undefined) delete process.env.FILE_CONTEXT_MAX_UPLOAD_BYTES; else process.env.FILE_CONTEXT_MAX_UPLOAD_BYTES = previous;
    }
  });

  it("rejects an oversized image response mid-stream instead of materializing it", async () => {
    vi.resetModules();
    const previousTargets = process.env.IMAGE_ALLOWED_TARGETS;
    const previousFetch = globalThis.fetch;
    process.env.IMAGE_ALLOWED_TARGETS = "https://example.com";
    let bytesProduced = 0;
    globalThis.fetch = async () => ({
      ok: true,
      status: 200,
      headers: new Map([["content-type", "image/png"]]),
      // No content-length: the only defense left is the streaming bound.
      body: (async function* stream() {
        for (let index = 0; index < 4096; index += 1) {
          bytesProduced += 64 * 1024;
          yield Buffer.alloc(64 * 1024, 1);
        }
      })(),
    });
    try {
      const { fetchImageAsset } = await import("../server/image-asset.mjs");
      await expect(fetchImageAsset("https://example.com/big.png")).rejects.toThrow(/too large/i);
      // Proves the read stopped early rather than consuming the full 256 MB.
      expect(bytesProduced).toBeLessThan(32 * 1024 * 1024);
    } finally {
      globalThis.fetch = previousFetch;
      if (previousTargets === undefined) delete process.env.IMAGE_ALLOWED_TARGETS; else process.env.IMAGE_ALLOWED_TARGETS = previousTargets;
    }
  });
});

describe("outbound timeouts", () => {
  it("applies an abort signal to both image-search requests", async () => {
    const previousFetch = globalThis.fetch;
    const seen = [];
    globalThis.fetch = async (input, init = {}) => {
      seen.push({ url: String(input), signal: init.signal, headerSignal: init.headers?.signal });
      if (seen.length === 1) {
        return { ok: true, status: 200, headers: new Map([["content-type", "text/html"]]), body: (async function* stream() { yield Buffer.from("vqd=\"12345\""); })() };
      }
      return { ok: true, status: 200, headers: new Map([["content-type", "application/json"]]), body: (async function* stream() { yield Buffer.from(JSON.stringify({ results: [{ title: "t", image: "https://images.example.com/a.png" }] })); })() };
    };
    try {
      const results = await duckDuckGoImageSearch("test query", 4);
      expect(results).toHaveLength(1);
      expect(seen).toHaveLength(2);
      for (const call of seen) {
        // The signal must be a top-level fetch option. It was previously nested
        // inside headers on the second call, which silently disabled the timeout.
        expect(call.signal, `missing top-level signal for ${call.url}`).toBeInstanceOf(AbortSignal);
        expect(call.headerSignal).toBeUndefined();
      }
    } finally {
      globalThis.fetch = previousFetch;
    }
  });
});
