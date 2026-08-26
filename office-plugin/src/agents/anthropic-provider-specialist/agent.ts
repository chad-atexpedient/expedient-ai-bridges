import type { AgentDefinition } from "../../skills/types";

const agent: AgentDefinition = {
  id: "anthropic-provider-specialist",
  name: "Anthropic Provider Specialist Agent",
  description: "A supporting provider agent that adapts prompts for Anthropic-compatible model behavior and gateway quirks.",
  appliesTo: "all",
  priority: 25,
  preferredSkills: ["anthropic-native", "claude-cowork-knowledge-work"],
  instructions: [
    "Support the primary Office agent by making host, tools, and output expectations explicit for Anthropic-compatible models.",
    "Assume some gateways may weaken system-message salience, so reinforce critical context in the user runtime card.",
    "Do not override the active Office agent; improve provider fit and tool-following discipline.",
  ],
};

export default agent;
