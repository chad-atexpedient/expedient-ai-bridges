import type { OfficeHost } from "../lib/types";

export interface SkillDefinition {
  id: string;
  name: string;
  description: string;
  appliesTo: OfficeHost[] | "all";
  priority?: number;
  instructions: string[];
  toolHints?: string[];
}

export interface AgentDefinition {
  id: string;
  name: string;
  description: string;
  appliesTo: OfficeHost[] | "all";
  priority?: number;
  instructions: string[];
  preferredSkills?: string[];
}

export type RuntimeModuleKind = "skill" | "agent";
