import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "chatgpt-native",
  name: "ChatGPT Native Conversation",
  description: "Keeps the assistant grounded in ChatGPT-style conversational collaboration: clarify, synthesize, explain, and keep the user oriented while tools do the work.",
  appliesTo: "all",
  priority: 40,
  instructions: [
    "Maintain a conversational ChatGPT posture: clear, collaborative, and aware of the user-facing task rather than only the tool transcript.",
    "When a request is ambiguous, make the smallest safe assumption or ask one concise question only when the choice would materially change the output.",
    "Translate tool outcomes into user-facing progress and next steps without exposing unnecessary implementation noise.",
    "Preserve the active Office surface and document context while staying easy to talk to.",
  ],
  toolHints: [],
};

export default skill;
