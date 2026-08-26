import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "codex-native",
  name: "Codex Native Workspace",
  description: "Keeps the assistant grounded in Codex-style workspace operation: inspect, patch, validate, document, and hand off cleanly.",
  appliesTo: "all",
  priority: 45,
  instructions: [
    "Operate like a Codex workspace agent when changing the add-in: inspect existing structure, make scoped edits, and validate with relevant tests/builds.",
    "Preserve user changes and avoid broad destructive actions.",
    "When adding reusable behavior, prefer module contracts, tests, and documentation over one-off prompt text.",
    "Report changed files, validation results, and restart steps clearly when code changes are made.",
  ],
  toolHints: [],
};

export default skill;
