import { describe, expect, it } from "vitest";
import { agentsForHost, listAgents, listSkills, runtimeInstructionBundle, skillsForHost, supportingAgentsForHost } from "../src/skills/registry";
import { buildOpenAIRequest } from "../src/providers/adapters";
import type { ProviderSettings } from "../src/lib/types";

const base: ProviderSettings = {
  provider: "openai-compatible",
  apiKey: "test-key",
  baseUrl: "https://example.test/v1/",
  route: "openai-v1",
  modelMode: "auto",
  model: "GPT-5-Mini",
  autoModels: ["GPT-5-Mini", "GPT-5.4", "GPT-5.5"],
  temperature: null,
  maxTokens: null,
  useLocalProxy: true,
};

const messages = [{ id: "1", role: "user" as const, content: "Make a chart", timestamp: "now" }];

describe("skills and agents registry", () => {
  it("discovers live skill and agent modules through Vite glob imports", () => {
    expect(listSkills().map((skill) => skill.id)).toEqual(expect.arrayContaining(["excel-operator", "powerpoint-builder", "research-context", "office-generation", "ctrl-demo-showcase", "chatgpt-native", "codex-native", "anthropic-native", "claude-cowork-knowledge-work", "popular-agent-patterns"]));
    expect(listAgents().map((agent) => agent.id)).toEqual(expect.arrayContaining(["excel-analyst", "powerpoint-deck-builder", "office-generalist", "ctrl-demo-guide", "chatgpt-collaborator", "codex-workspace-engineer", "anthropic-provider-specialist", "claude-cowork-collaborator"]));
  });

  it("selects host-specific agent and skills", () => {
    expect(agentsForHost("excel")[0].id).toBe("excel-analyst");
    expect(skillsForHost("powerpoint").map((skill) => skill.id)).toEqual(expect.arrayContaining(["powerpoint-builder", "research-context", "office-generation", "ctrl-demo-showcase"]));
    const bundle = runtimeInstructionBundle("excel");
    expect(bundle.prompt).toContain("Active agent: Excel Analyst Agent");
    expect(bundle.prompt).toContain("Skill Excel Operator");
    expect(bundle.prompt).toContain("Supporting agents: Codex Workspace Engineer Agent");
    expect(supportingAgentsForHost("excel").map((agent) => agent.id)).toEqual(expect.arrayContaining(["chatgpt-collaborator", "codex-workspace-engineer", "ctrl-demo-guide", "anthropic-provider-specialist", "claude-cowork-collaborator"]));
  });

  it("injects active agent and skill context into provider prompts", () => {
    const request = buildOpenAIRequest(base, messages, { host: "excel", title: "Book.xlsx", selectionLabel: "Sheet1!A1", text: "", metadata: {} });
    const body = JSON.parse(String(request.init.body));
    expect(body.messages.at(-1).content).toContain("Active agent: Excel Analyst Agent");
    expect(body.messages.at(-1).content).toContain("Excel Operator");
    expect(body.messages.at(-1).content).toContain("Supporting agents: Codex Workspace Engineer Agent");
    expect(body.messages.at(-1).content).toContain("ChatGPT Collaborator Agent");
  });

  it("adds agent and skill instructions to hidden context when context/capabilities are present", () => {
    const request = buildOpenAIRequest(base, messages, { host: "excel", title: "Book.xlsx", selectionLabel: "Sheet1!A1", text: "Revenue", metadata: {} });
    const body = JSON.parse(String(request.init.body));
    expect(body.messages[1].content).toContain("Agent and skill runtime instructions");
    expect(body.messages[1].content).toContain("Excel Analyst Agent");
    expect(body.messages[1].content).toContain("ChatGPT Native Conversation");
    expect(body.messages[1].content).toContain("Codex Native Workspace");
    expect(body.messages[1].content).toContain("Anthropic Native Reasoning");
    expect(body.messages[1].content).toContain("Claude Cowork Knowledge Work");
  });
});
