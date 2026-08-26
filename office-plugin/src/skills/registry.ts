import type { OfficeHost } from "../lib/types";
import type { AgentDefinition, SkillDefinition } from "./types";

const skillModules = import.meta.glob<{ default?: SkillDefinition; skill?: SkillDefinition }>("./**/skill.{ts,js}", { eager: true });
const agentModules = import.meta.glob<{ default?: AgentDefinition; agent?: AgentDefinition }>("../agents/**/agent.{ts,js}", { eager: true });

function byPriorityThenName<T extends { priority?: number; name: string }>(a: T, b: T) {
  return (b.priority ?? 0) - (a.priority ?? 0) || a.name.localeCompare(b.name);
}

function appliesToHost(appliesTo: OfficeHost[] | "all", host: OfficeHost) {
  return appliesTo === "all" || appliesTo.includes(host);
}

export function listSkills() {
  return Object.values(skillModules)
    .map((module) => module.default ?? module.skill)
    .filter((skill): skill is SkillDefinition => Boolean(skill?.id && skill?.instructions?.length))
    .sort(byPriorityThenName);
}

export function listAgents() {
  return Object.values(agentModules)
    .map((module) => module.default ?? module.agent)
    .filter((agent): agent is AgentDefinition => Boolean(agent?.id && agent?.instructions?.length))
    .sort(byPriorityThenName);
}

export function skillsForHost(host: OfficeHost) {
  return listSkills().filter((skill) => appliesToHost(skill.appliesTo, host));
}

export function agentsForHost(host: OfficeHost) {
  return listAgents().filter((agent) => appliesToHost(agent.appliesTo, host));
}

export function primaryAgentsForHost(host: OfficeHost) {
  const agents = listAgents();
  const hostSpecific = agents.filter((agent) => agent.appliesTo !== "all" && appliesToHost(agent.appliesTo, host));
  return (hostSpecific.length ? hostSpecific : agents.filter((agent) => agent.appliesTo === "all")).sort(byPriorityThenName);
}

export function supportingAgentsForHost(host: OfficeHost) {
  const primary = primaryAgentsForHost(host)[0];
  return agentsForHost(host).filter((agent) => agent.id !== primary?.id && agent.appliesTo === "all");
}

export function selectAgentForHost(host: OfficeHost) {
  return primaryAgentsForHost(host)[0] ?? null;
}

export function runtimeInstructionBundle(host: OfficeHost) {
  const agent = selectAgentForHost(host);
  const supportingAgents = supportingAgentsForHost(host);
  const skills = skillsForHost(host);
  const lines: string[] = [];
  if (agent) {
    lines.push(`Active agent: ${agent.name} (${agent.id})`);
    lines.push(agent.description);
    lines.push(...agent.instructions.map((instruction) => `Agent rule: ${instruction}`));
  }
  if (supportingAgents.length) {
    lines.push(`Supporting agents: ${supportingAgents.map((supportingAgent) => `${supportingAgent.name} (${supportingAgent.id})`).join(", ")}`);
    for (const supportingAgent of supportingAgents) {
      lines.push(`Supporting agent ${supportingAgent.name}: ${supportingAgent.description}`);
      lines.push(...supportingAgent.instructions.map((instruction) => `Supporting agent rule: ${instruction}`));
    }
  }
  if (skills.length) {
    lines.push(`Active skills: ${skills.map((skill) => `${skill.name} (${skill.id})`).join(", ")}`);
    for (const skill of skills) {
      lines.push(`Skill ${skill.name}: ${skill.description}`);
      lines.push(...skill.instructions.map((instruction) => `Skill rule: ${instruction}`));
      if (skill.toolHints?.length) lines.push(`Skill tool hints: ${skill.toolHints.join(", ")}`);
    }
  }
  return { agent, supportingAgents, skills, prompt: lines.join("\n") };
}

export type { AgentDefinition, SkillDefinition } from "./types";
