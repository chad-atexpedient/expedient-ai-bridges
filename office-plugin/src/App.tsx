import { Bot, ChevronDown, ChevronUp, FileInput, MessageSquarePlus, Paperclip, PanelRightOpen, Send, Settings, Sparkle, Sparkles, X } from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { SettingsPanel } from "./components/SettingsPanel";
import { DEFAULT_SETTINGS } from "./lib/defaults";
import { summarizeAttachment } from "./lib/artifacts";
import { clearSharedSettings, loadSettings, loadSharedSettings, saveSharedSettings } from "./lib/storage";
import type { AppSettings, AttachmentContext, ChatMessage, DocumentContext, OfficeCapabilitySnapshot, OfficeHost, ToolCallRequest, ToolCallResult } from "./lib/types";
import { blockedToolResult, detectOfficeCapabilities } from "./lib/capabilities";
import { executeToolCall } from "./lib/tools";
import { addSessionMemory, createChatMemory, loadSessionMemories, memoryToAttachment, rankChatMemories, saveSessionMemories, type ChatMemoryRecord } from "./lib/sessionMemory";
import { attachmentToUploadedAsset, registerUploadedAsset, unregisterUploadedAsset } from "./lib/uploadRegistry";
import { officeReady, readDocumentContext, safeReadDocumentContext, insertIntoOffice } from "./office/host";
import { completeChat, testProvider } from "./providers/adapters";
import "./styles/app.css";

function createMessage(role: ChatMessage["role"], content: string): ChatMessage {
  return { id: crypto.randomUUID(), role, content, timestamp: new Date().toISOString() };
}

function hostLabel(host: OfficeHost) {
  return host === "excel" ? "Excel" : host === "word" ? "Word" : host === "powerpoint" ? "PowerPoint" : "Preview";
}

const TEXT_ATTACHMENT_TYPES = new Set(["text/plain", "text/csv", "text/tab-separated-values", "application/json", "text/markdown"]);
const TEXT_ATTACHMENT_EXTENSIONS = /\.(txt|csv|tsv|json|md|markdown|log)$/i;
const MAX_ATTACHMENT_BYTES = 25_000_000;

function isImageFile(file: File) {
  return file.type.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp)$/i.test(file.name);
}

function isPowerPointTemplateFile(file: File) {
  return /\.pptx$/i.test(file.name) || file.type === "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}

function dataUrlParts(dataUrl: string) {
  const comma = dataUrl.indexOf(",");
  return comma >= 0 ? { header: dataUrl.slice(0, comma), base64: dataUrl.slice(comma + 1) } : { header: "", base64: dataUrl };
}

function readFileAsDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error(`${file.name} could not be read.`));
    reader.readAsDataURL(file);
  });
}

function canReadAsText(file: File) {
  return TEXT_ATTACHMENT_TYPES.has(file.type) || TEXT_ATTACHMENT_EXTENSIONS.test(file.name);
}

async function attachmentFromFile(file: File): Promise<AttachmentContext> {
  if (file.size > MAX_ATTACHMENT_BYTES) throw new Error(`${file.name} is too large for the context reader. Use a file under 25 MB.`);
  if (isImageFile(file)) {
    const dataUrl = await readFileAsDataUrl(file);
    const { base64 } = dataUrlParts(dataUrl);
    return {
      id: crypto.randomUUID(),
      name: file.name,
      type: file.type || "image/*",
      size: file.size,
      text: `Uploaded image asset: ${file.name} (${file.type || "image"}, ${file.size.toLocaleString()} bytes). Native Office image tools can insert this by assetId or filename.`,
      assetKind: "image",
      dataUrl,
      base64,
    };
  }
  if (!canReadAsText(file)) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = "";
    const chunkSize = 0x8000;
    for (let offset = 0; offset < bytes.length; offset += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
    }
    const base64 = btoa(binary);
    const response = await fetch("/api/file-context", {
      method: "POST",
      headers: { "content-type": "application/json", accept: "application/json" },
      body: JSON.stringify({ name: file.name, type: file.type || "application/octet-stream", size: file.size, base64, maxChars: 18000 }),
    });
    const json = await response.json().catch(() => null);
    if (!response.ok) throw new Error(json?.error?.message || `${file.name} could not be read.`);
    const isTemplate = isPowerPointTemplateFile(file);
    const type = json.type || file.type || "application/octet-stream";
    return {
      id: crypto.randomUUID(),
      name: file.name,
      type,
      size: json.size || file.size,
      text: [
        `Uploaded file context extracted with strategy: ${json.strategy || "unknown"}`,
        isTemplate ? "This uploaded PowerPoint file is also available as a template asset. Generated deck tools can reference it by attachment id or filename." : "",
        json.brandProfile ? `Extracted brand profile: ${JSON.stringify(json.brandProfile)}` : "",
        json.clipped ? "The file context was clipped to fit the model context window." : "",
        json.extractedText || "",
      ].filter(Boolean).join("\n"),
      assetKind: isTemplate ? "office-template" : undefined,
      dataUrl: isTemplate ? `data:${type};base64,${base64}` : undefined,
      base64: isTemplate ? base64 : undefined,
      brandProfile: json.brandProfile,
    };
  }
  if (file.size > 2_000_000) throw new Error(`${file.name} is too large for direct text reading. Use a file under 2 MB or a supported Office/binary upload under 25 MB.`);
  return {
    id: crypto.randomUUID(),
    name: file.name,
    type: file.type || "text/plain",
    size: file.size,
    text: await file.text(),
  };
}

interface TaskItem {
  id: string;
  name: string;
  status: "running" | "done" | "failed";
  detail: string;
}

function friendlyToolName(name: string) {
  return name.replace(/_/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

interface DemoCardData {
  kind: "ctrl-demo-showcase";
  title: string;
  subtitle: string;
  host: OfficeHost;
  features: Array<{ title: string; detail: string }>;
  actions: Array<{ id: "live" | "artifact"; label: string; detail: string }>;
}

export function isSimpleDemoRequest(text: string) {
  const normalized = text.toLowerCase();
  const asksForDemo = /\b(show me a demo|demo|showcase|tour|what can (this|you|ctrl) do|features?)\b/.test(normalized);
  const asksForMutation = /\b(create|make|build|insert|generate|download|artifact|file|deck|slide|sheet|worksheet|workbook|document|docx|pptx|xlsx)\b/.test(normalized);
  return asksForDemo && !asksForMutation;
}

export function demoCardFor(host: OfficeHost, productName: string): DemoCardData {
  const label = hostLabel(host);
  const surfaceLine = host === "unknown" ? "I'm in preview mode, so this is a UI tour of what the add-in can do once opened in Office." : `I'm connected to ${label}, so I can use this chat pane as the control center for the open file.`;
  return {
    kind: "ctrl-demo-showcase",
    title: `${productName} demo`,
    subtitle: surfaceLine,
    host,
    features: [
      { title: "Chat with the file", detail: "Read the current selection or document context, then answer with awareness of where you are." },
      { title: "Use tools while it thinks", detail: "Search the web, read uploaded files, and call Office actions while the task tray shows progress." },
      { title: "Act in Office when asked", detail: host === "powerpoint" ? "Create slides, text boxes, tables, shapes, backgrounds, and images in the current deck." : host === "excel" ? "Write ranges, create tables/charts, format sheets, validate cells, and add comments." : host === "word" ? "Draft sections, add headings/tables/comments/images, and format the document." : "Open in Excel, Word, or PowerPoint to unlock live Office actions." },
      { title: "BYOK + shared settings", detail: "Use an OpenAI-compatible or Anthropic-compatible key, shared branding, model routing, and file context." },
    ],
    actions: [
      { id: "live", label: host === "unknown" ? "Open in Office first" : `Try live in ${label}`, detail: host === "unknown" ? "Live Office demos need Excel, Word, or PowerPoint." : "Creates a small live sample in the current file." },
      { id: "artifact", label: "Make sample files", detail: "Creates downloadable Excel, Word, and PowerPoint examples." },
    ],
  };
}

export function encodeDemoCard(card: DemoCardData) {
  return `::ctrl-demo-showcase\n${JSON.stringify(card)}\n::`;
}

export function parseDemoCard(content: string): DemoCardData | null {
  const match = content.match(/^::ctrl-demo-showcase\n([\s\S]+)\n::$/);
  if (!match) return null;
  try {
    const parsed = JSON.parse(match[1]);
    return parsed?.kind === "ctrl-demo-showcase" ? parsed : null;
  } catch {
    return null;
  }
}


function ThinkingSparkles() {
  return (
    <span className="thinking-sparkles" aria-hidden="true">
      <Sparkle className="sparkle-main" size={20} />
      <Sparkle className="sparkle-dot dot-one" size={7} />
      <Sparkle className="sparkle-dot dot-two" size={5} />
      <Sparkle className="sparkle-dot dot-three" size={8} />
    </span>
  );
}

export default function App() {
  const [settings, setSettings] = useState<AppSettings>(() => loadSettings());
  const [host, setHost] = useState<OfficeHost>("unknown");
  const [context, setContext] = useState<DocumentContext | null>(null);
  const [capabilities, setCapabilities] = useState<OfficeCapabilitySnapshot | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>([
    createMessage("assistant", "I can read the current Office context, answer with your own OpenAI-compatible or Anthropic-compatible key, and insert the result back into the file."),
  ]);
  const [draft, setDraft] = useState("Summarize the selected content and suggest the next useful edit.");
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isBusy, setIsBusy] = useState(false);
  const [testStatus, setTestStatus] = useState("");
  const [notice, setNotice] = useState("");
  const [retryText, setRetryText] = useState("");
  const [writeReview, setWriteReview] = useState<{ host: OfficeHost; text: string } | null>(null);
  const [isResetReviewOpen, setIsResetReviewOpen] = useState(false);
  const [attachments, setAttachments] = useState<AttachmentContext[]>([]);
  const [tasks, setTasks] = useState<TaskItem[]>([]);
  const [isTaskListOpen, setIsTaskListOpen] = useState(false);
  const [streamingMessageId, setStreamingMessageId] = useState<string | null>(null);
  const [workStatus, setWorkStatus] = useState("Idle");
  const [workStartedAt, setWorkStartedAt] = useState<number | null>(null);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);
  const [chatMemories, setChatMemories] = useState<ChatMemoryRecord[]>(() => loadSessionMemories());
  const [activeMemoryTitles, setActiveMemoryTitles] = useState<string[]>([]);
  const requestControllerRef = useRef<AbortController | null>(null);
  const settingsTriggerRef = useRef<HTMLButtonElement>(null);
  const settingsDialogRef = useRef<HTMLElement>(null);

  useEffect(() => {
    document.documentElement.style.setProperty("--brand-primary", settings.branding.primaryColor);
    document.documentElement.style.setProperty("--brand-accent", settings.branding.accentColor);
    document.title = settings.branding.productName;
  }, [settings]);

  useEffect(() => {
    if (!isSettingsOpen) {
      settingsTriggerRef.current?.focus();
      return;
    }
    const dialog = settingsDialogRef.current;
    dialog?.querySelector<HTMLElement>("input, select, textarea, button")?.focus();
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setIsSettingsOpen(false);
      if (event.key !== "Tab") return;
      const focusable = Array.from(dialog?.querySelectorAll<HTMLElement>("button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [href], [tabindex]:not([tabindex='-1'])") || []);
      if (focusable.length < 2) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [isSettingsOpen]);

  useEffect(() => {
    let cancelled = false;
    loadSharedSettings().then((sharedSettings) => {
      if (!cancelled) setSettings(sharedSettings);
    });
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    officeReady().then(async (readyHost) => {
      setHost(readyHost);
      const nextCapabilities = detectOfficeCapabilities(readyHost);
      setCapabilities(nextCapabilities);
      const result = await safeReadDocumentContext(readyHost);
      setContext(result.context);
      if (result.warning) setNotice(result.warning);
    });
  }, []);

  useEffect(() => {
    if (!workStartedAt || !isBusy) return;
    const timer = window.setInterval(() => setElapsedSeconds(Math.floor((Date.now() - workStartedAt) / 1000)), 1000);
    return () => window.clearInterval(timer);
  }, [isBusy, workStartedAt]);

  useEffect(() => {
    saveSessionMemories(chatMemories);
  }, [chatMemories]);

  const lastAssistantMessage = useMemo(() => [...messages].reverse().find((message) => message.role === "assistant" && message.content), [messages]);
  const taskSummary = useMemo(() => {
    const last = tasks.at(-1);
    return {
      count: tasks.length,
      lastName: last ? friendlyToolName(last.name) : workStatus,
      running: tasks.some((task) => task.status === "running") || isBusy,
    };
  }, [tasks, workStatus, isBusy]);

  const runTool = async (call: ToolCallRequest) => {
    const blocked = capabilities ? blockedToolResult(call, capabilities) : null;
    if (blocked) return blocked;
    return executeToolCall(host, call);
  };

  const rememberCurrentChat = () => {
    const memory = createChatMemory(messages, host, context);
    if (!memory) return null;
    setChatMemories((current) => addSessionMemory(current, memory));
    return memory;
  };

  const startNewChat = () => {
    const memory = rememberCurrentChat();
    setMessages([
      createMessage("assistant", memory
        ? `Started a new chat. I kept a compact memory of "${memory.title}" so I can pull it back in if it becomes relevant.`
        : "Started a new chat. I can still use the current Office context and attached files."),
    ]);
    setDraft("What should we work on next?");
    setTasks([]);
    setActiveMemoryTitles([]);
    setNotice(memory ? `Cached prior chat: ${memory.title}` : "New chat started.");
  };

  const refreshContext = async () => {
    setNotice("Reading current Office context...");
    try {
      const nextContext = await readDocumentContext(host);
      setContext(nextContext);
      setNotice(`Context refreshed from ${nextContext.selectionLabel}.`);
    } catch (error: any) {
      setNotice(error?.message || String(error));
    }
  };

  const addAssistantMessage = (content: string) => {
    setMessages((current) => [...current, createMessage("assistant", content)]);
  };

  const runDemoAction = async (mode: "live" | "artifact") => {
    if (isBusy) return;
    if (mode === "live" && host === "unknown") {
      setNotice("Open the add-in in Excel, Word, or PowerPoint to try a live Office demo.");
      return;
    }
    setIsBusy(true);
    setTasks([]);
    setIsTaskListOpen(false);
    setWorkStatus(mode === "live" ? "Creating live demo..." : "Creating sample files...");
    setWorkStartedAt(Date.now());
    try {
      const call: ToolCallRequest = {
        id: crypto.randomUUID(),
        name: "ctrl_create_demo_showcase",
        arguments: mode === "live" ? { surface: "current", mode: "live" } : { surface: "all", mode: "artifact" },
      };
      setTasks([{ id: call.id, name: call.name, status: "running", detail: mode === "live" ? "Creating a small live demo in the current file..." : "Creating downloadable sample files..." }]);
      const result = await runTool(call);
      setTasks((current) => current.map((task) => task.id === result.id ? { ...task, status: result.ok ? "done" : "failed", detail: result.content.slice(0, 260) } : task));
      if (!result.ok) {
        addAssistantMessage(`I could not run that demo action: ${result.content}`);
        return;
      }
      const parsed = JSON.parse(result.content || "{}");
      if (mode === "live") {
        addAssistantMessage(parsed.message || `Created a live ${hostLabel(host)} demo in the current file.`);
        return;
      }
      const links = Array.isArray(parsed.generated)
        ? parsed.generated.map((item: any) => item?.result?.markdownLink || item?.result?.downloadUrl).filter(Boolean)
        : [];
      addAssistantMessage(links.length ? `Created sample files:\n\n${links.join("\n")}` : "Created the sample files.");
    } catch (error: any) {
      addAssistantMessage(`I could not run that demo action: ${error?.message || String(error)}`);
    } finally {
      setIsBusy(false);
      setWorkStatus("Idle");
      setWorkStartedAt(null);
    }
  };

  const sendMessage = async (messageOverride?: string) => {
    const text = (messageOverride ?? draft).trim();
    if (!text || isBusy) return;

    setIsBusy(true);
    const controller = new AbortController();
    requestControllerRef.current = controller;
    setNotice("");
    setRetryText("");
    setTasks([]);
    setIsTaskListOpen(false);
    setWorkStatus("Preparing request...");
    setWorkStartedAt(Date.now());
    setElapsedSeconds(0);

    const userMessage = createMessage("user", text);
    const nextMessages = [...messages, userMessage];
    setMessages(nextMessages);
    setDraft("");

    if (isSimpleDemoRequest(text)) {
      setMessages((current) => [...current, createMessage("assistant", encodeDemoCard(demoCardFor(host, settings.branding.productName)))]);
      setIsBusy(false);
      setWorkStatus("Idle");
      setWorkStartedAt(null);
      return;
    }

    try {
      const currentCapabilities = detectOfficeCapabilities(host);
      setCapabilities(currentCapabilities);
      let currentContext: DocumentContext | null = null;
      let contextWarning = "";
      if (settings.includeDocumentContext) {
        setWorkStatus("Reading Office context...");
        const result = await safeReadDocumentContext(host);
        currentContext = result.context;
        contextWarning = result.warning;
        setContext(result.context);
        if (result.warning) setNotice(result.warning);
      }

      const relevantMemories = rankChatMemories(`${text}\n${currentContext?.text ?? ""}`, chatMemories, 3);
      setActiveMemoryTitles(relevantMemories.map((memory) => memory.title));
      const memoryAttachments = relevantMemories.map(memoryToAttachment);

      let assistantMessage: ChatMessage | null = null;
      let streamedText = "";
      const result = await completeChat({
        settings: settings.provider,
        messages: nextMessages,
        context: currentContext,
        attachments: [...memoryAttachments, ...attachments],
        contextWarning,
        capabilities: currentCapabilities,
        toolsEnabled: true,
        stream: true,
        signal: controller.signal,
        toolExecutor: async (call) => {
          const blocked = blockedToolResult(call, currentCapabilities);
          if (blocked) return blocked;
          return executeToolCall(host, call);
        },
        onStatus: setWorkStatus,
        onToken: (token) => {
          streamedText += token;
          if (!assistantMessage) {
            assistantMessage = createMessage("assistant", streamedText);
            setStreamingMessageId(assistantMessage.id);
            setMessages((current) => [...current, assistantMessage!]);
            return;
          }
          setMessages((current) => current.map((message) => message.id === assistantMessage!.id ? { ...message, content: streamedText } : message));
        },
        onToolStart: (call) => {
          setIsTaskListOpen(false);
          setTasks((current) => [...current, { id: call.id, name: call.name, status: "running", detail: "Running..." }]);
        },
        onToolResult: (result: ToolCallResult) => {
          setTasks((current) => current.map((task) => task.id === result.id ? { ...task, status: result.ok ? "done" : "failed", detail: result.content.slice(0, 260) } : task));
        },
      });

      if (!streamedText) setMessages((current) => [...current, createMessage("assistant", result.text)]);
    } catch (error: any) {
      if (error?.name === "AbortError" || controller.signal.aborted) {
        setNotice("Request cancelled. Your message remains in the conversation; you can send it again when ready.");
      } else {
        setRetryText(text);
        setNotice(`Request failed: ${error?.message || String(error)}`);
        setMessages((current) => [...current, createMessage("assistant", `I could not complete that request: ${error?.message || String(error)}`)]);
      }
    } finally {
      if (requestControllerRef.current === controller) requestControllerRef.current = null;
      setIsBusy(false);
      setStreamingMessageId(null);
      setWorkStatus("Idle");
      setWorkStartedAt(null);
    }
  };

  const runTest = async (candidate: AppSettings) => {
    setTestStatus("Testing...");
    try {
      const result = await testProvider(candidate.provider);
      setTestStatus(`Connected: ${result.text}`);
      return true;
    } catch (error: any) {
      setTestStatus(`Connection failed: ${error?.message || String(error)}`);
      return false;
    }
  };

  const insertLast = async () => {
    if (!lastAssistantMessage) return;
    setWriteReview({ host, text: lastAssistantMessage.content });
  };

  const confirmInsert = async () => {
    if (!writeReview) return;
    const reviewed = writeReview;
    setWriteReview(null);
    setNotice("Inserting response...");
    try {
      setNotice(await insertIntoOffice(reviewed.host, reviewed.text));
    } catch (error: any) {
      setNotice(error?.message || String(error));
    }
  };

  const attachFiles = async (files?: FileList | null) => {
    if (!files?.length) return;
    const next: AttachmentContext[] = [];
    const failures: string[] = [];
    for (const file of Array.from(files)) {
      try {
        next.push(await attachmentFromFile(file));
      } catch (error: any) {
        failures.push(error?.message || String(error));
      }
    }
    if (next.length) {
      for (const attachment of next) {
        const asset = attachmentToUploadedAsset(attachment);
        if (asset) registerUploadedAsset(asset);
      }
      setAttachments((current) => [...current, ...next].slice(-6));
    }
    setNotice([
      next.length ? `Attached ${next.length} file${next.length === 1 ? "" : "s"} for chat context.` : "",
      ...failures,
    ].filter(Boolean).join(" "));
  };

  const resetSettings = async () => {
    setSettings(DEFAULT_SETTINGS);
    await clearSharedSettings();
    await saveSharedSettings(DEFAULT_SETTINGS);
    setNotice("Shared settings reset.");
  };

  const saveSettingsAndClose = async (nextSettings: AppSettings) => {
    const saved = await saveSharedSettings(nextSettings);
    setSettings(saved);
    setIsSettingsOpen(false);
    setNotice("Shared settings saved for Excel, Word, and PowerPoint.");
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true">
            {settings.branding.logoDataUrl ? <img src={settings.branding.logoDataUrl} alt="" /> : <Sparkles size={22} />}
          </div>
          <div>
            <h1>{settings.branding.productName}</h1>
            <p>{hostLabel(host)} BYOK add-in</p>
          </div>
        </div>
        <button ref={settingsTriggerRef} className="icon-button" aria-label="Open settings" title="Settings" onClick={() => setIsSettingsOpen((open) => !open)}>
          <Settings size={18} />
        </button>
      </header>

      <section className="context-strip">
        <div>
          <span>{context?.selectionLabel || "No context loaded"}</span>
          <small>
            {context?.text ? `${context.text.length.toLocaleString()} characters available` : "Open in Office or use preview mode"}
            {context?.title ? ` | Source: ${context.title}` : ""}
            {chatMemories.length ? ` | ${chatMemories.length} prior chat memor${chatMemories.length === 1 ? "y" : "ies"} cached` : ""}
          </small>
          <div className="provenance-line" aria-label="Current grounding sources">
            {context?.title ? <span>Office: {context.title}</span> : <span>Office context: not loaded</span>}
            {attachments.length > 0 && <span>Attachments: {attachments.map((item) => item.name).join(", ")}</span>}
            {activeMemoryTitles.length > 0 && <span>Memory: {activeMemoryTitles.join(", ")}</span>}
          </div>
        </div>
        <div className="context-actions">
          <button className="secondary-button compact" onClick={startNewChat} disabled={isBusy}>
            <MessageSquarePlus size={15} />
            New chat
          </button>
          <button className="secondary-button compact" onClick={refreshContext}>
            <FileInput size={15} />
            Read context
          </button>
        </div>
      </section>
      {host === "unknown" && settings.deploymentMode === "preview" && (
        <section className="setup-note" aria-label="Preview setup guidance">
          <strong>Preview mode</strong>
          <span>Explore the task-pane workflow without sending context to a provider. Open Settings to choose Personal BYOK or Managed Enterprise mode.</span>
          <button className="secondary-button compact" type="button" onClick={() => setIsSettingsOpen(true)}>Open setup</button>
        </section>
      )}
      <section className="trust-note" role="status">
        <strong>{settings.deploymentMode === "managed" ? "Managed enterprise mode" : settings.deploymentMode === "personal" ? "Personal BYOK mode" : "Preview mode"}</strong>
        <span>{settings.deploymentMode === "managed" ? "Provider access and retention follow your organization’s policy." : settings.deploymentMode === "personal" ? "Your provider key is kept in this Office profile; review the provider URL before sending context." : "Connect a provider in Settings when you are ready to test the workflow."}</span>
      </section>

      {activeMemoryTitles.length > 0 && (
        <div className="memory-strip">
          Grounded with prior chat memory: {activeMemoryTitles.join("; ")}
        </div>
      )}

      {notice && <div className="notice" role="status" aria-live="polite">{notice}{retryText && <button className="notice-action" type="button" onClick={() => { void sendMessage(retryText); }}>Retry</button>}</div>}
      {(tasks.length > 0 || isBusy) && (
        <section className={`task-tray ${isTaskListOpen ? "open" : "collapsed"}`} aria-label="Task progress">
          <button className="task-summary" type="button" onClick={() => setIsTaskListOpen((open) => !open)}>
            {isBusy && <ThinkingSparkles />}
            <span>{taskSummary.count} tool call{taskSummary.count === 1 ? "" : "s"}</span>
            <small>{isBusy ? `${workStatus} | ${elapsedSeconds}s` : `${taskSummary.running ? "Running" : "Last"}: ${taskSummary.lastName}`}</small>
            {isTaskListOpen ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
          </button>
          {isBusy && <button type="button" className="cancel-button" onClick={() => requestControllerRef.current?.abort()}>Cancel</button>}
          {isTaskListOpen && (
            <div className="task-list">
              {tasks.length === 0 && <p>{workStatus} | {elapsedSeconds}s elapsed</p>}
              {tasks.map((task) => (
                <article className={`task-item ${task.status}`} key={task.id}>
                  <span>{task.status === "running" ? "..." : task.status === "done" ? "OK" : "!"}</span>
                  <div>
                    <strong>{friendlyToolName(task.name)}</strong>
                    <small>{task.detail}</small>
                  </div>
                </article>
              ))}
            </div>
          )}
        </section>
      )}

      <div className="workspace chat-workspace">
        <section className="chat-panel" aria-label="Assistant chat">
          <div className="messages">
            {messages.map((message) => (
              <article key={message.id} className={`message ${message.role}`}>
                <div className="avatar">{message.role === "assistant" ? <Bot size={15} /> : "You"}</div>
                {message.role === "assistant" && parseDemoCard(message.content) ? (() => {
                  const card = parseDemoCard(message.content)!;
                  return (
                    <div className="demo-card">
                      <div className="demo-card-hero">
                        <span><Sparkles size={18} /></span>
                        <div>
                          <strong>{card.title}</strong>
                          <small>{card.subtitle}</small>
                        </div>
                      </div>
                      <div className="demo-feature-grid">
                        {card.features.map((feature) => (
                          <section key={feature.title}>
                            <strong>{feature.title}</strong>
                            <span>{feature.detail}</span>
                          </section>
                        ))}
                      </div>
                      <div className="demo-actions">
                        {card.actions.map((action) => (
                          <button
                            key={action.id}
                            type="button"
                            className={action.id === "live" ? "primary-button" : "secondary-button"}
                            onClick={() => runDemoAction(action.id)}
                            disabled={isBusy || (action.id === "live" && card.host === "unknown")}
                            title={action.detail}
                          >
                            {action.id === "live" ? <Sparkle size={15} /> : <FileInput size={15} />}
                            {action.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })() : <p>{message.content}</p>}
              </article>
            ))}
            {isBusy && !streamingMessageId && (
              <article className="message assistant">
                <div className="avatar"><Bot size={15} /></div>
                <p className="thinking-line"><ThinkingSparkles /> {workStatus} <span aria-hidden="true">·</span> <span>{elapsedSeconds}s</span></p>
              </article>
            )}
          </div>

          <div className="composer">
            {attachments.length > 0 && (
              <div className="attachment-shelf" aria-label="Attached files">
                {attachments.map((attachment) => (
                  <span className="attachment-chip" key={attachment.id} title={summarizeAttachment(attachment.name, attachment.text, 800)}>
                    <Paperclip size={13} />
                    {attachment.name}
                    <button type="button" title={`Remove ${attachment.name}`} onClick={() => { unregisterUploadedAsset(attachment.id); setAttachments((current) => current.filter((item) => item.id !== attachment.id)); }}>
                      <X size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
            <textarea aria-label="Message" value={draft} onChange={(event) => setDraft(event.target.value)} placeholder="Ask about the selected cells, text, deck, or attached files..." />
            <div className="composer-actions">
              <label className="secondary-button attach-button">
                <Paperclip size={16} />
                Attach file
                <input type="file" multiple accept=".txt,.csv,.tsv,.json,.md,.markdown,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.pdf,.png,.jpg,.jpeg,.gif,.webp,.bmp,text/*,image/*,application/json,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(event) => attachFiles(event.target.files)} />
              </label>
              <button className="secondary-button" onClick={insertLast} disabled={!lastAssistantMessage}>
                <PanelRightOpen size={16} />
                Insert last
              </button>
              <button className="primary-button" onClick={() => { void sendMessage(); }} disabled={!draft.trim() || isBusy}>
                <Send size={16} />
                Send
              </button>
            </div>
          </div>
        </section>

        {isSettingsOpen && (
          <SettingsPanel
            dialogRef={settingsDialogRef}
            settings={settings}
            onSave={saveSettingsAndClose}
            onCancel={() => setIsSettingsOpen(false)}
            onTest={runTest}
            testStatus={testStatus}
          />
        )}
      </div>

      <footer className="footer-note">
        {settings.deploymentMode === "managed" ? "Managed mode uses organization-controlled provider access; personal keys are not stored here." : "Personal BYOK keys stay in local Office web storage. Review the provider and context before sending."}
        <button onClick={() => setIsResetReviewOpen(true)}>Reset all</button>
      </footer>

      {isResetReviewOpen && (
        <div className="review-backdrop" role="presentation">
          <section className="write-review" role="dialog" aria-modal="true" aria-labelledby="reset-review-title">
            <div className="panel-heading">
              <div>
                <h2 id="reset-review-title">Reset local settings?</h2>
                <p>This clears the saved provider, branding, deployment mode, and shared settings from this add-in profile. It does not revoke provider or Entra sessions outside CTRL.</p>
              </div>
              <button className="icon-button" type="button" aria-label="Close reset confirmation" onClick={() => setIsResetReviewOpen(false)}><X size={16} /></button>
            </div>
            <div className="settings-actions review-actions">
              <button className="secondary-button" type="button" onClick={() => setIsResetReviewOpen(false)}>Cancel</button>
              <button className="primary-button" type="button" onClick={() => { setIsResetReviewOpen(false); void resetSettings(); }}>Reset settings</button>
            </div>
          </section>
        </div>
      )}

      {writeReview && (
        <div className="review-backdrop" role="presentation">
          <section className="write-review" role="dialog" aria-modal="true" aria-labelledby="write-review-title">
            <div className="panel-heading">
              <div>
                <h2 id="write-review-title">Review Office change</h2>
                <p>CTRL is ready to write the following response into {hostLabel(writeReview.host)}.</p>
              </div>
              <button className="icon-button" type="button" aria-label="Close write review" onClick={() => setWriteReview(null)}><X size={16} /></button>
            </div>
            <div className="write-review-meta"><strong>Target</strong><span>Active {hostLabel(writeReview.host)} file</span><strong>Content</strong><span>{writeReview.text.length.toLocaleString()} characters from the latest assistant response</span></div>
            <pre className="write-review-preview">{writeReview.text.slice(0, 4000)}{writeReview.text.length > 4000 ? "\n…" : ""}</pre>
            <div className="settings-actions review-actions">
              <button className="secondary-button" type="button" onClick={() => setWriteReview(null)}>Cancel</button>
              <button className="primary-button" type="button" onClick={() => { void confirmInsert(); }}>Write to file</button>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
