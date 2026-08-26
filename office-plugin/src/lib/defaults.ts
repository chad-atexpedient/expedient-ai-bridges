import type { AppSettings } from "./types";

export const RECOMMENDED_AUTO_MODELS: [string, string, string] = ["GPT-5-Mini", "GPT-5.4", "GPT-5.5"];

export const DEFAULT_SETTINGS: AppSettings = {
  deploymentMode: "preview",
  provider: {
    provider: "openai-compatible",
    apiKey: "",
    baseUrl: "https://api.openai.com/v1",
    route: "openai-v1",
    modelMode: "auto",
    model: "GPT-5-Mini",
    autoModels: RECOMMENDED_AUTO_MODELS,
    temperature: null,
    maxTokens: null,
    useLocalProxy: true,
  },
  branding: {
    productName: "Expedient AI Workspace",
    primaryColor: "#2458d3",
    accentColor: "#11a37f",
    logoDataUrl: null,
  },
  includeDocumentContext: true,
};
