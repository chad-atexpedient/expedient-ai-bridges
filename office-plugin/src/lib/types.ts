export type OfficeHost = "excel" | "word" | "powerpoint" | "unknown";

export type ProviderKind = "openai-compatible" | "anthropic-compatible";

export type ProviderRoute = "openai-v1" | "openwebui-api" | "anthropic-v1" | "custom";

export type ModelMode = "auto" | "manual";

export type DeploymentMode = "preview" | "personal" | "managed";

export interface ProviderSettings {
  provider: ProviderKind;
  apiKey: string;
  baseUrl: string;
  route: ProviderRoute;
  modelMode: ModelMode;
  model: string;
  autoModels: [string, string, string];
  temperature: number | null;
  maxTokens: number | null;
  useLocalProxy: boolean;
}

export interface BrandingSettings {
  productName: string;
  primaryColor: string;
  accentColor: string;
  logoDataUrl: string | null;
}

export interface AppSettings {
  deploymentMode: DeploymentMode;
  provider: ProviderSettings;
  branding: BrandingSettings;
  includeDocumentContext: boolean;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  timestamp: string;
}

export interface DocumentContext {
  host: OfficeHost;
  title: string;
  selectionLabel: string;
  text: string;
  metadata: Record<string, unknown>;
}

export interface AttachmentContext {
  id: string;
  name: string;
  type: string;
  size: number;
  text: string;
  assetKind?: "image" | "office-template";
  dataUrl?: string;
  base64?: string;
  brandProfile?: PowerPointBrandProfile;
}

export interface PowerPointBrandProfile {
  source: string | null;
  colors: Record<string, string>;
  fonts: { major: string; minor: string };
  layouts: Array<{ index: number; name: string; type: string; placeholders: string[] }>;
  media: Array<{ name: string; type: string; bytes: number; candidate: boolean }>;
  guidance: string[];
}

export interface ToolCallRequest {
  id: string;
  name: string;
  arguments: Record<string, unknown>;
}

export interface ToolCallResult {
  id: string;
  name: string;
  ok: boolean;
  content: string;
}

export interface OfficeCapabilitySnapshot {
  host: OfficeHost;
  platform: string;
  requirementSets: Record<string, boolean>;
  availableTools: string[];
  unavailableTools: Record<string, string>;
  notes: string[];
}

export type ToolExecutor = (call: ToolCallRequest) => Promise<ToolCallResult>;

export interface CompletionRequest {
  settings: ProviderSettings;
  messages: ChatMessage[];
  context?: DocumentContext | null;
  attachments?: AttachmentContext[];
  contextWarning?: string;
  capabilities?: OfficeCapabilitySnapshot | null;
  toolsEnabled?: boolean;
  stream?: boolean;
  signal?: AbortSignal;
  toolExecutor?: ToolExecutor;
  onToken?: (token: string) => void;
  onStatus?: (status: string) => void;
  onToolStart?: (call: ToolCallRequest) => void;
  onToolResult?: (result: ToolCallResult) => void;
}

export interface CompletionResult {
  text: string;
  raw?: unknown;
  toolResults?: ToolCallResult[];
}

