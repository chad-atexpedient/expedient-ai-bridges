import type { AttachmentContext, ChatMessage, DocumentContext, OfficeHost } from "./types";

export interface ChatMemoryRecord {
  id: string;
  title: string;
  summary: string;
  keywords: string[];
  host: OfficeHost;
  documentTitle: string;
  messageCount: number;
  createdAt: string;
  updatedAt: string;
}

const SESSION_MEMORY_KEY = "ctrl-byok-office-addin:session-memory:v1";
const MAX_MEMORIES = 24;
const MAX_SUMMARY_CHARS = 3600;
const STOPWORDS = new Set([
  "about", "after", "again", "also", "and", "are", "because", "been", "but", "can", "could", "did", "does", "for", "from", "have", "here", "into", "like", "more", "need", "that", "the", "their", "then", "there", "this", "through", "use", "using", "was", "were", "what", "when", "where", "which", "with", "would", "you", "your",
]);

function safeRandomId() {
  return globalThis.crypto?.randomUUID?.() ?? `memory-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export function tokenizeMemoryText(text: string) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9_\-\s]/g, " ")
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 2 && !STOPWORDS.has(token));
}

function topKeywords(text: string, limit = 14) {
  const counts = new Map<string, number>();
  for (const token of tokenizeMemoryText(text)) counts.set(token, (counts.get(token) ?? 0) + 1);
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, limit)
    .map(([token]) => token);
}

function shortLine(text: string, limit = 180) {
  return text.replace(/\s+/g, " ").trim().slice(0, limit);
}

export function createChatMemory(messages: ChatMessage[], host: OfficeHost, context?: DocumentContext | null): ChatMemoryRecord | null {
  const conversation = messages.filter((message) => message.role === "user" || message.role === "assistant");
  const userMessages = conversation.filter((message) => message.role === "user");
  if (!userMessages.length) return null;

  const firstUser = userMessages[0];
  const title = shortLine(firstUser.content, 72) || "Prior chat";
  const compactLines = conversation.slice(-12).map((message) => `${message.role === "user" ? "User" : "Assistant"}: ${shortLine(message.content, 320)}`);
  const summary = [
    `Prior chat: ${title}`,
    `Host: ${host}`,
    context?.title ? `Document: ${context.title}` : "",
    context?.selectionLabel ? `Last known Office context: ${context.selectionLabel}` : "",
    "Compact transcript:",
    ...compactLines,
  ].filter(Boolean).join("\n").slice(0, MAX_SUMMARY_CHARS);
  const now = new Date().toISOString();

  return {
    id: safeRandomId(),
    title,
    summary,
    keywords: topKeywords(`${title}\n${summary}\n${context?.text ?? ""}`),
    host,
    documentTitle: context?.title || "Untitled Office file",
    messageCount: conversation.length,
    createdAt: now,
    updatedAt: now,
  };
}

export function rankChatMemories(query: string, memories: ChatMemoryRecord[], limit = 3) {
  const queryTokens = new Set(tokenizeMemoryText(query));
  if (!queryTokens.size) return memories.slice(0, limit);

  return memories
    .map((memory) => {
      const memoryTokens = new Set([...memory.keywords, ...tokenizeMemoryText(`${memory.title}\n${memory.summary}`)]);
      let score = 0;
      for (const token of queryTokens) {
        if (memoryTokens.has(token)) score += memory.keywords.includes(token) ? 3 : 1;
        if (memory.title.toLowerCase().includes(token)) score += 2;
      }
      return { memory, score };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || b.memory.updatedAt.localeCompare(a.memory.updatedAt))
    .slice(0, limit)
    .map((item) => item.memory);
}

export function memoryToAttachment(memory: ChatMemoryRecord): AttachmentContext {
  return {
    id: `memory-${memory.id}`,
    name: `Prior chat memory - ${memory.title}`,
    type: "application/x-ctrl-chat-memory",
    size: memory.summary.length,
    text: [
      memory.summary,
      memory.keywords.length ? `Keywords: ${memory.keywords.join(", ")}` : "",
      `Memory created: ${memory.createdAt}`,
    ].filter(Boolean).join("\n"),
  };
}

export function loadSessionMemories(storage: Storage = window.sessionStorage): ChatMemoryRecord[] {
  try {
    const raw = storage.getItem(SESSION_MEMORY_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((item) => item?.id && item?.summary).slice(0, MAX_MEMORIES);
  } catch {
    return [];
  }
}

export function saveSessionMemories(memories: ChatMemoryRecord[], storage: Storage = window.sessionStorage) {
  storage.setItem(SESSION_MEMORY_KEY, JSON.stringify(memories.slice(0, MAX_MEMORIES)));
}

export function addSessionMemory(memories: ChatMemoryRecord[], memory: ChatMemoryRecord) {
  return [memory, ...memories.filter((item) => item.id !== memory.id)].slice(0, MAX_MEMORIES);
}

export { SESSION_MEMORY_KEY };
