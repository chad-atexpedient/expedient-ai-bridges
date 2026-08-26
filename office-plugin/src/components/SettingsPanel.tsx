import { RotateCcw, Save, Upload, Wifi, X } from "lucide-react";
import { useEffect, useState, type RefObject } from "react";
import { RECOMMENDED_AUTO_MODELS } from "../lib/defaults";
import type { AppSettings, DeploymentMode, ProviderKind, ProviderRoute } from "../lib/types";

interface Props {
  settings: AppSettings;
  onSave(settings: AppSettings): void | Promise<void>;
  onCancel(): void;
  onTest(settings: AppSettings): Promise<boolean>;
  testStatus: string;
  dialogRef?: RefObject<HTMLElement | null>;
}

function providerDefaults(provider: ProviderKind) {
  return provider === "openai-compatible"
    ? { baseUrl: "https://api.openai.com/v1", route: "openai-v1" as ProviderRoute, model: RECOMMENDED_AUTO_MODELS[0] }
    : { baseUrl: "https://api.anthropic.com/v1", route: "anthropic-v1" as ProviderRoute, model: "claude-3-5-sonnet-latest" };
}

function numberOrNull(value: string) {
  return value.trim() === "" ? null : Number(value);
}

export function SettingsPanel({ settings, onSave, onCancel, onTest, testStatus, dialogRef }: Props) {
  const [draft, setDraft] = useState<AppSettings>(settings);
  const [isTesting, setIsTesting] = useState(false);
  const [validationMessage, setValidationMessage] = useState("");

  useEffect(() => setDraft(settings), [settings]);

  const update = (patch: Partial<AppSettings>) => setDraft((current) => ({ ...current, ...patch }));
  const updateProvider = (patch: Partial<AppSettings["provider"]>) => update({ provider: { ...draft.provider, ...patch } });
  const updateBranding = (patch: Partial<AppSettings["branding"]>) => update({ branding: { ...draft.branding, ...patch } });

  const handleLogo = (file?: File) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setValidationMessage("Choose a PNG, JPG, SVG, or other image file for the logo.");
      return;
    }
    if (file.size > 1_000_000) {
      setValidationMessage("Choose a logo under 1 MB so it can be saved locally in the add-in.");
      return;
    }
    setValidationMessage("");
    const reader = new FileReader();
    reader.onload = () => updateBranding({ logoDataUrl: String(reader.result) });
    reader.readAsDataURL(file);
  };

  const testAndMaybeSave = async (saveAfterTest: boolean) => {
    setValidationMessage("");
    setIsTesting(true);
    const ok = await onTest(draft);
    setIsTesting(false);
    if (ok && saveAfterTest) await onSave(draft);
  };

  return (
    <div className="settings-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}>
      <section ref={dialogRef} className="settings-drawer" aria-labelledby="settings-title" role="dialog" aria-modal="true">
        <div className="panel-heading sticky-heading">
          <div>
            <h2 id="settings-title">BYOK settings</h2>
            <p>Test, save, then return to chat.</p>
            {draft.deploymentMode !== settings.deploymentMode && (
              <p className="status-text" role="status" aria-live="polite">Mode preview only — save to apply it to the task pane.</p>
            )}
            {validationMessage && <p className="error-text" role="alert">{validationMessage}</p>}
          </div>
          <button className="icon-button" title="Close settings" onClick={onCancel}>
            <X size={16} />
          </button>
        </div>

        <div className="settings-scroll">
          <label>
            Deployment mode
            <select value={draft.deploymentMode} onChange={(event) => update({ deploymentMode: event.target.value as DeploymentMode })}>
              <option value="preview">Preview — no provider key saved</option>
              <option value="personal">Personal BYOK — stored in this Office profile</option>
              <option value="managed">Managed enterprise — organization policy and vault</option>
            </select>
            <span className="status-text" role="status" aria-live="polite">{draft.deploymentMode === "managed" ? "Your organization controls providers, storage, retention, and access." : draft.deploymentMode === "personal" ? "Your key stays in this add-in profile and is sent only to the configured provider." : "Preview mode shows the workflow without requiring a provider connection."}</span>
          </label>

          <label>
            Provider
            <select
              value={draft.provider.provider}
              onChange={(event) => {
                const provider = event.target.value as ProviderKind;
                updateProvider({ provider, ...providerDefaults(provider) });
              }}
            >
              <option value="openai-compatible">OpenAI compatible</option>
              <option value="anthropic-compatible">Anthropic compatible</option>
            </select>
          </label>

          <label>
            API key
            <input
              type="password"
              autoComplete="off"
              value={draft.provider.apiKey}
              placeholder={draft.deploymentMode === "managed" ? "Configured by your organization" : "Paste your key here"}
              disabled={draft.deploymentMode === "managed"}
              onChange={(event) => updateProvider({ apiKey: event.target.value })}
            />
            <span className="status-text">{draft.deploymentMode === "managed" ? "Managed mode never stores a personal provider key in this browser or settings service." : "Personal mode keeps this key in the local Office profile; review the provider URL before sending context."}</span>
          </label>

          <label>
            Base URL
            <input value={draft.provider.baseUrl} onChange={(event) => updateProvider({ baseUrl: event.target.value })} />
          </label>

          <label>
            Route
            <select value={draft.provider.route} onChange={(event) => updateProvider({ route: event.target.value as ProviderRoute })}>
              <option value="openai-v1">OpenAI-compatible base URL adds /chat/completions</option>
              <option value="openwebui-api">OpenWebUI API base URL adds /chat/completions</option>
              <option value="anthropic-v1">Anthropic-compatible base URL adds /messages</option>
              <option value="custom">Custom full chat/messages URL</option>
            </select>
          </label>

          <label>
            Model routing
            <select value={draft.provider.modelMode} onChange={(event) => updateProvider({ modelMode: event.target.value as "auto" | "manual" })}>
              <option value="auto">Auto route by complexity</option>
              <option value="manual">Manual model selection</option>
            </select>
          </label>

          {draft.provider.modelMode === "auto" ? (
            <div className="model-grid">
              {draft.provider.autoModels.map((model, index) => (
                <label key={index}>
                  {index === 0 ? "Light" : index === 1 ? "Standard" : "Deep"}
                  <input
                    value={model}
                    onChange={(event) => {
                      const next = [...draft.provider.autoModels] as [string, string, string];
                      next[index] = event.target.value;
                      updateProvider({ autoModels: next });
                    }}
                  />
                </label>
              ))}
            </div>
          ) : (
            <label>
              Model
              <select value={draft.provider.model} onChange={(event) => updateProvider({ model: event.target.value })}>
                {draft.provider.autoModels.map((model) => <option key={model} value={model}>{model}</option>)}
                <option value={draft.provider.model}>{draft.provider.model}</option>
              </select>
              <input value={draft.provider.model} onChange={(event) => updateProvider({ model: event.target.value })} />
            </label>
          )}

          <div className="two-col">
            <label>
              Temperature
              <input
                type="number"
                min="0"
                max="2"
                step="0.1"
                placeholder="Gateway default"
                value={draft.provider.temperature ?? ""}
                onChange={(event) => updateProvider({ temperature: numberOrNull(event.target.value) })}
              />
            </label>
            <label>
              Max tokens
              <input
                type="number"
                min="1"
                max="32000"
                placeholder="Gateway default"
                value={draft.provider.maxTokens ?? ""}
                onChange={(event) => updateProvider({ maxTokens: numberOrNull(event.target.value) })}
              />
            </label>
          </div>
          <p className="status-text">Leave temperature and max tokens blank to let the model gateway choose supported defaults.</p>

          <label className="check-row">
            <input type="checkbox" checked={draft.provider.useLocalProxy} onChange={(event) => updateProvider({ useLocalProxy: event.target.checked })} />
            Use local proxy for OpenWebUI/LiteLLM CORS
          </label>

          <div className="divider" />

          <label>
            Add-in name
            <input value={draft.branding.productName} onChange={(event) => updateBranding({ productName: event.target.value })} />
          </label>

          <div className="two-col">
            <label>
              Primary
              <input type="color" value={draft.branding.primaryColor} onChange={(event) => updateBranding({ primaryColor: event.target.value })} />
            </label>
            <label>
              Accent
              <input type="color" value={draft.branding.accentColor} onChange={(event) => updateBranding({ accentColor: event.target.value })} />
            </label>
          </div>

          <label className="upload-control">
            <Upload size={16} />
            Upload branded logo
            <input type="file" accept="image/*" onChange={(event) => handleLogo(event.target.files?.[0])} />
          </label>

          <button className="secondary-button" title="Reset uploaded logo" onClick={() => updateBranding({ logoDataUrl: null })}>
            <RotateCcw size={16} />
            Reset logo
          </button>

          <label className="check-row">
            <input type="checkbox" checked={draft.includeDocumentContext} onChange={(event) => update({ includeDocumentContext: event.target.checked })} />
            Include current Office context with prompts
          </label>
        </div>

        <div className="settings-actions">
          {testStatus && <p className="status-text settings-status" role="status" aria-live="polite">{testStatus}</p>}
          <button className="secondary-button" onClick={() => testAndMaybeSave(false)} disabled={isTesting}>
            <Wifi size={16} />
            {isTesting ? "Testing..." : "Test"}
          </button>
          <button className="primary-button" onClick={() => testAndMaybeSave(true)} disabled={isTesting}>
            <Save size={16} />
            Test and save
          </button>
        </div>
      </section>
    </div>
  );
}
