import { describe, expect, it } from "vitest";
import { addSessionMemory, createChatMemory, memoryToAttachment, rankChatMemories } from "../src/lib/sessionMemory";
import type { ChatMessage } from "../src/lib/types";

const messages: ChatMessage[] = [
  { id: "a", role: "assistant", content: "Ready.", timestamp: "2026-01-01T00:00:00Z" },
  { id: "u1", role: "user", content: "Use the FY27 planning notes to build revenue assumptions", timestamp: "2026-01-01T00:01:00Z" },
  { id: "a1", role: "assistant", content: "I created a revenue assumptions table using 12 percent growth and noted the source.", timestamp: "2026-01-01T00:02:00Z" },
];

describe("session chat memory", () => {
  it("compacts a prior chat into a reusable memory record", () => {
    const memory = createChatMemory(messages, "excel", { host: "excel", title: "Model.xlsx", selectionLabel: "Sheet1!A1", text: "Revenue 12 percent growth", metadata: {} });
    expect(memory?.title).toContain("FY27 planning notes");
    expect(memory?.summary).toContain("Compact transcript");
    expect(memory?.keywords).toContain("revenue");
  });

  it("ranks relevant memories and turns them into hidden context attachments", () => {
    const revenue = createChatMemory(messages, "excel", null)!;
    const unrelated = createChatMemory([
      { id: "u2", role: "user", content: "Draft a welcome email", timestamp: "now" },
      { id: "a2", role: "assistant", content: "Email drafted.", timestamp: "now" },
    ], "word", null)!;
    const memories = addSessionMemory(addSessionMemory([], unrelated), revenue);
    const ranked = rankChatMemories("continue the revenue model assumptions", memories, 1);
    expect(ranked[0].id).toBe(revenue.id);
    const attachment = memoryToAttachment(ranked[0]);
    expect(attachment.type).toBe("application/x-ctrl-chat-memory");
    expect(attachment.text).toContain("revenue assumptions");
  });
});
