import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "anthropic-native",
  name: "Anthropic Native Reasoning",
  description: "Keeps Anthropic-compatible models grounded in explicit context, careful tool use, and provider-shim tolerance.",
  appliesTo: "all",
  priority: 43,
  instructions: [
    "Repeat the active Office surface and task constraints plainly because Anthropic-compatible gateways may handle system context differently.",
    "Use explicit step intent before tool use: what is being read, what will be changed, and what success should look like.",
    "Prefer tool-backed changes over prose when native tools are available; if tools are unavailable, explain the limitation clearly.",
    "Keep final answers concise and concrete: completed actions, source assumptions, and any remaining gap.",
  ],
  toolHints: [],
};

export default skill;
