import { z } from "zod";
import { DEFAULT_SETTINGS, RECOMMENDED_AUTO_MODELS } from "./defaults";
import type { AppSettings } from "./types";

const STORAGE_KEY = "expedient-ai-workspace:v2";
const LEGACY_STORAGE_KEY = "ctrl-byok-office-addin:v1";
const SETTINGS_API_PATH = "/api/settings";

const nullableNumber = z.union([z.number(), z.null()]).default(null);

const SettingsSchema = z.object({
  deploymentMode: z.enum(["preview", "personal", "managed"]).default("preview"),
  provider: z.object({
    provider: z.enum(["openai-compatible", "anthropic-compatible"]),
    apiKey: z.string(),
    baseUrl: z.string().url(),
    route: z.enum(["openai-v1", "openwebui-api", "anthropic-v1", "custom"]).default("openai-v1"),
    modelMode: z.enum(["auto", "manual"]).default("auto"),
    model: z.string().min(1),
    autoModels: z.tuple([z.string().min(1), z.string().min(1), z.string().min(1)]).default(RECOMMENDED_AUTO_MODELS),
    temperature: nullableNumber.refine((value) => value === null || (value >= 0 && value <= 2), "Temperature must be blank or between 0 and 2"),
    maxTokens: nullableNumber.refine((value) => value === null || (Number.isInteger(value) && value >= 1 && value <= 32000), "Max tokens must be blank or between 1 and 32000"),
    useLocalProxy: z.boolean().default(true),
  }),
  branding: z.object({
    productName: z.string().min(1).max(48),
    primaryColor: z.string().regex(/^#[0-9a-fA-F]{6}$/),
    accentColor: z.string().regex(/^#[0-9a-fA-F]{6}$/),
    logoDataUrl: z.string().nullable(),
  }),
  includeDocumentContext: z.boolean(),
});

function normalizeParsedSettings(parsed: any) {
  return {
    ...DEFAULT_SETTINGS,
    ...parsed,
    provider: {
      ...DEFAULT_SETTINGS.provider,
      ...parsed.provider,
      autoModels: parsed.provider?.autoModels ?? DEFAULT_SETTINGS.provider.autoModels,
      temperature: parsed.provider?.temperature ?? null,
      maxTokens: parsed.provider?.maxTokens ?? null,
    },
    branding: { ...DEFAULT_SETTINGS.branding, ...parsed.branding },
  };
}

export function parseSettings(value: unknown): AppSettings {
  return SettingsSchema.parse(normalizeParsedSettings(value));
}

export function loadSettings(storage: Storage = window.localStorage): AppSettings {
  let raw = storage.getItem(STORAGE_KEY);
  if (!raw) {
    const legacy = storage.getItem(LEGACY_STORAGE_KEY);
    if (legacy) {
      try {
        const migrated = parseSettings(JSON.parse(legacy));
        const safe = migrated.deploymentMode === "managed" ? { ...migrated, provider: { ...migrated.provider, apiKey: "" } } : migrated;
        storage.setItem(STORAGE_KEY, JSON.stringify(safe));
        raw = storage.getItem(STORAGE_KEY);
      } catch { /* Preserve legacy state for rollback. */ }
    }
  }
  if (!raw) return DEFAULT_SETTINGS;

  try {
    return parseSettings(JSON.parse(raw));
  } catch {
    return DEFAULT_SETTINGS;
  }
}

export function saveSettings(settings: AppSettings, storage: Storage = window.localStorage) {
  const parsed = parseSettings(settings);
  const persisted = parsed.deploymentMode === "managed"
    ? { ...parsed, provider: { ...parsed.provider, apiKey: "" } }
    : parsed;
  storage.setItem(STORAGE_KEY, JSON.stringify(persisted));
}

export function clearSettings(storage: Storage = window.localStorage) {
  storage.removeItem(STORAGE_KEY);
}

async function parseJsonResponse(response: Response) {
  if (response.status === 204) return null;
  const json = await response.json();
  if (!response.ok) {
    const message = json?.error?.message || json?.message || `${response.status} ${response.statusText}`;
    throw new Error(message);
  }
  return json;
}

export async function loadSharedSettings(storage: Storage = window.localStorage, fetcher: typeof fetch = fetch): Promise<AppSettings> {
  try {
    const response = await fetcher(SETTINGS_API_PATH, { headers: { accept: "application/json" } });
    const json = await parseJsonResponse(response);
    if (!json) return loadSettings(storage);
    const settings = parseSettings(json.settings ?? json);
    saveSettings(settings, storage);
    return settings;
  } catch {
    return loadSettings(storage);
  }
}

export async function saveSharedSettings(settings: AppSettings, storage: Storage = window.localStorage, fetcher: typeof fetch = fetch): Promise<AppSettings> {
  const parsed = parseSettings(settings);
  const localSafe = parsed.deploymentMode === "managed"
    ? { ...parsed, provider: { ...parsed.provider, apiKey: "" } }
    : parsed;
  saveSettings(localSafe, storage);
  try {
    const response = await fetcher(SETTINGS_API_PATH, {
      method: "PUT",
      headers: { "content-type": "application/json", accept: "application/json" },
      body: JSON.stringify({ settings: localSafe }),
    });
    const json = await parseJsonResponse(response);
    if (!json) return localSafe;
    const saved = parseSettings(json.settings ?? json);
    saveSettings(saved, storage);
    return saved;
  } catch {
    return localSafe;
  }
}

export async function clearSharedSettings(storage: Storage = window.localStorage, fetcher: typeof fetch = fetch) {
  clearSettings(storage);
  try {
    await fetcher(SETTINGS_API_PATH, { method: "DELETE" });
  } catch {
    // Local fallback is already cleared.
  }
}

export { SETTINGS_API_PATH, STORAGE_KEY };
