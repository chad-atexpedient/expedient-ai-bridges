import type { AgentDefinition } from "../../skills/types";

const agent: AgentDefinition = {
  id: "office-generalist",
  name: "Office Generalist Agent",
  description: "Fallback Office operator for preview or unknown host mode.",
  appliesTo: "all",
  priority: 1,
  preferredSkills: ["research-context", "office-generation"],
  instructions: [
    "If the Office host is unknown, avoid claiming live document edits succeeded.",
    "Use structured output or generated Office files when live Office context is unavailable.",
  ],
};

export default agent;
