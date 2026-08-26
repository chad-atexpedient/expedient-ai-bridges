import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS } from "../src/lib/defaults";
import { loadSettings, loadSharedSettings, saveSettings, saveSharedSettings, SETTINGS_API_PATH } from "../src/lib/storage";

class MemoryStorage implements Storage {
  private values = new Map<string, string>();
  get length() { return this.values.size; }
  clear() { this.values.clear(); }
  getItem(key: string) { return this.values.get(key) ?? null; }
  key(index: number) { return [...this.values.keys()][index] ?? null; }
  removeItem(key: string) { this.values.delete(key); }
  setItem(key: string, value: string) { this.values.set(key, value); }
}

describe("settings storage", () => {
  it("persists BYOK settings and branding", () => {
    const storage = new MemoryStorage();
    const settings = {
      ...DEFAULT_SETTINGS,
      provider: { ...DEFAULT_SETTINGS.provider, apiKey: "abc", baseUrl: "https://llm.example/v1", route: "openai-v1" as const, modelMode: "auto" as const, autoModels: ["GPT-5-Mini", "GPT-5.4", "GPT-5.5"] as [string, string, string], temperature: null, maxTokens: null, useLocalProxy: true },
      branding: { ...DEFAULT_SETTINGS.branding, productName: "Acme AI", logoDataUrl: "data:image/png;base64,abc" },
    };
    saveSettings(settings, storage);
    expect(loadSettings(storage).provider.apiKey).toBe("abc");
    expect(loadSettings(storage).branding.productName).toBe("Acme AI");
  });

  it("falls back to defaults for invalid settings", () => {
    const storage = new MemoryStorage();
    storage.setItem("ctrl-byok-office-addin:v1", "{bad json");
    expect(loadSettings(storage)).toEqual(DEFAULT_SETTINGS);
  });

  it("loads shared settings from the same-origin settings endpoint and mirrors them locally", async () => {
    const storage = new MemoryStorage();
    const shared = {
      ...DEFAULT_SETTINGS,
      provider: { ...DEFAULT_SETTINGS.provider, apiKey: "shared-key", baseUrl: "https://shared.example/v1" },
      branding: { ...DEFAULT_SETTINGS.branding, productName: "Shared CTRL" },
    };

    const loaded = await loadSharedSettings(storage, async (input, init) => {
      expect(input).toBe(SETTINGS_API_PATH);
      expect(init?.headers).toEqual({ accept: "application/json" });
      return new Response(JSON.stringify({ settings: shared }), { status: 200, headers: { "content-type": "application/json" } });
    });

    expect(loaded.provider.apiKey).toBe("shared-key");
    expect(loadSettings(storage).branding.productName).toBe("Shared CTRL");
  });

  it("saves shared settings through the endpoint and keeps a local fallback mirror", async () => {
    const storage = new MemoryStorage();
    const settings = {
      ...DEFAULT_SETTINGS,
      provider: { ...DEFAULT_SETTINGS.provider, apiKey: "saved-key", baseUrl: "https://saved.example/v1" },
      branding: { ...DEFAULT_SETTINGS.branding, productName: "Saved CTRL" },
    };

    const saved = await saveSharedSettings(settings, storage, async (input, init) => {
      expect(input).toBe(SETTINGS_API_PATH);
      expect(init?.method).toBe("PUT");
      const body = JSON.parse(String(init?.body));
      expect(body.settings.provider.apiKey).toBe("saved-key");
      return new Response(JSON.stringify({ settings: body.settings }), { status: 200, headers: { "content-type": "application/json" } });
    });

    expect(saved.branding.productName).toBe("Saved CTRL");
    expect(loadSettings(storage).provider.apiKey).toBe("saved-key");
  });

  it("falls back to the local mirror when the shared settings endpoint is unavailable", async () => {
    const storage = new MemoryStorage();
    saveSettings({ ...DEFAULT_SETTINGS, branding: { ...DEFAULT_SETTINGS.branding, productName: "Local CTRL" } }, storage);

    const loaded = await loadSharedSettings(storage, async () => {
      throw new Error("offline");
    });

    expect(loaded.branding.productName).toBe("Local CTRL");
  });

  it("does not persist a provider key in managed enterprise mode", () => {
    const storage = new MemoryStorage();
    saveSettings({
      ...DEFAULT_SETTINGS,
      deploymentMode: "managed",
      provider: { ...DEFAULT_SETTINGS.provider, apiKey: "must-not-persist" },
    }, storage);
    expect(loadSettings(storage).provider.apiKey).toBe("");
    expect(storage.getItem("ctrl-byok-office-addin:v1")).not.toContain("must-not-persist");
  });
});




