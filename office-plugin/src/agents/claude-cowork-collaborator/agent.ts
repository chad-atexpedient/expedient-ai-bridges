import type { AgentDefinition } from "../../skills/types";

const agent: AgentDefinition = {
  id: "claude-cowork-collaborator",
  name: "Claude Cowork Collaborator Agent",
  description: "A supporting cowork-style agent for business workflows, shared context, repeatable skill routing, and knowledge-work artifacts.",
  appliesTo: "all",
  priority: 24,
  preferredSkills: ["claude-cowork-knowledge-work", "popular-agent-patterns"],
  instructions: [
    "Support the primary Office agent by treating tasks as cowork-style knowledge workflows with reusable inputs, operations, and outputs.",
    "Favor context-rich, artifact-producing work over generic advice.",
    "When the user asks for a document, deck, or workbook, reinforce real Office artifact creation and no-placeholder behavior.",
  ],
};

export default agent;
