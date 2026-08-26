# Skills and agents runtime modules

The add-in now has a live-importable skills and agents structure under `src/`:

- `src/skills/<skill-id>/skill.ts` or `skill.js`
- `src/agents/<agent-id>/agent.ts` or `agent.js`

Vite discovers these modules through eager `import.meta.glob` calls in `src/skills/registry.ts`. During development, adding or editing one of these files is picked up by the normal Vite dev service without adding a central import by hand.

## Skill module contract

A skill exports a default object with this shape:

```ts
import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "example-skill",
  name: "Example Skill",
  description: "What this skill helps the active Office agent do.",
  appliesTo: ["excel"], // or "all"
  priority: 50,
  instructions: ["Concrete model behavior rule."],
  toolHints: ["excel_write_range"],
};

export default skill;
```

## Agent module contract

An agent exports a default object with this shape:

```ts
import type { AgentDefinition } from "../../skills/types";

const agent: AgentDefinition = {
  id: "example-agent",
  name: "Example Agent",
  description: "The operating role for a host or workflow.",
  appliesTo: ["powerpoint"], // or "all"
  priority: 100,
  preferredSkills: ["powerpoint-builder"],
  instructions: ["Concrete operating rule."],
};

export default agent;
```

## Runtime behavior

For each chat request, the provider adapter selects one primary agent for the current Office host and all skills that apply to that host. Host-specific agents win primary selection. All-host agents, such as ChatGPT and Codex support agents, are included as supporting agents so the model feels native to those spaces without overriding the active Excel, Word, or PowerPoint operator. Their instructions are injected in two places:

1. Hidden context: `Agent and skill runtime instructions`.
2. The latest user prompt runtime card: active agent, supporting agents, and active skills.

This redundancy is intentional for BYOK compatibility. Some gateways preserve system messages differently, so the model gets the active Office role in both hidden and user-visible prompt context.

## Current default modules

Skills:

- `excel-operator`
- `powerpoint-builder`
- `word-editor`
- `research-context`
- `office-generation`
- `ctrl-demo-showcase`
- `chatgpt-native`
- `codex-native`
- `anthropic-native`
- `claude-cowork-knowledge-work`
- `popular-agent-patterns`

Agents:

- `excel-analyst`
- `powerpoint-deck-builder`
- `word-document-specialist`
- `office-generalist`
- `ctrl-demo-guide`
- `chatgpt-collaborator`
- `codex-workspace-engineer`
- `anthropic-provider-specialist`
- `claude-cowork-collaborator`

## Prepared demo showcase

The demo layer gives the add-in a reliable answer to "show me what this can do" without depending on a model to invent a demo from scratch every time.

- Skill: `ctrl-demo-showcase`
- Supporting agent: `ctrl-demo-guide`
- Tool: `ctrl_create_demo_showcase`

Simple demo requests render directly in the chat pane as a lightweight add-in feature card. `ctrl_create_demo_showcase` is used only after the user asks for or clicks an explicit demo action: `surface=current` with `mode=live` adds a small live sample in the active Office file, while `surface=all` or `mode=artifact` creates safe downloadable Excel, Word, and PowerPoint showcase files. The artifact demos remain the richer fallback for generated chart parts, speaker notes, TOC/footnotes/revisions, and other Open XML features that Office.js may not expose live.

## Anthropic/Cowork and popular-agent research notes

The Anthropic/Cowork layer was based on public GitHub ecosystem patterns rather than copied code. Useful high-signal repositories found during research included:

- [anthropics/knowledge-work-plugins](https://github.com/anthropics/knowledge-work-plugins) - Apache-2.0, Claude Cowork knowledge-worker plugin patterns.
- [anthropics/claude-plugins-community](https://github.com/anthropics/claude-plugins-community) - Apache-2.0, community plugin marketplace signal.
- [OpenCoworkAI/open-cowork](https://github.com/OpenCoworkAI/open-cowork) - MIT, Cowork-style desktop app with Claude Code, MCP, sandbox, and skills concepts.
- [iOfficeAI/AionUi](https://github.com/iOfficeAI/AionUi) - Apache-2.0, high-star cowork/agent-team ecosystem signal across Claude Code, Codex, and other CLIs.
- [VoltAgent/awesome-agent-skills](https://github.com/VoltAgent/awesome-agent-skills) - MIT, high-star skills marketplace/curation pattern across Claude Code, Codex, Gemini, Cursor, and related agents.
- [affaan-m/ECC](https://github.com/affaan-m/ECC) - MIT, high-star agent harness signal around skills, memory, security, and research-first development.
- [BloopAI/vibe-kanban](https://github.com/BloopAI/vibe-kanban) - Apache-2.0, high-star agent workflow/kanban signal for coordinating coding agents.

Patterns imported into this add-in:

- skills as small modular recipes;
- host/provider-specific supporting agents instead of one monolithic prompt;
- context-first and MCP/M365/web-aware workflows;
- cowork-style knowledge-worker framing: inputs, operations, outputs, validation;
- explicit permission/runtime boundaries;
- final handoff that names what changed and how it was validated.
