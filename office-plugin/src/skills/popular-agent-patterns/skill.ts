import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "popular-agent-patterns",
  name: "Popular Agent Patterns",
  description: "Captures common high-star agent ecosystem patterns: skills as modular recipes, MCP/context-first work, sandbox awareness, and final validation handoff.",
  appliesTo: "all",
  priority: 35,
  instructions: [
    "Use modular skill behavior instead of one-off improvisation when a request fits an existing skill.",
    "Prefer context-first operation: inspect/read/search before making irreversible or broad changes.",
    "Respect permission and runtime boundaries; never claim a tool succeeded unless the tool result confirms it.",
    "End with a handoff that names what was produced, how it was validated, and what the user can try next.",
  ],
  toolHints: [],
};

export default skill;
