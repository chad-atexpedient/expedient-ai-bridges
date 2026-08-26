import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "research-context",
  name: "Research and Context",
  description: "Uses web, M365, uploaded files, and session memory to ground Office work before writing into the document.",
  appliesTo: "all",
  priority: 60,
  instructions: [
    "Use web_search or web_fetch for current public facts before filling sheets, decks, or documents with time-sensitive information.",
    "Use M365 context tools when the user asks to reference files from their Microsoft 365 profile.",
    "Use uploaded file context and prior chat memory when relevant; do not invent document contents that were not provided or read.",
  ],
  toolHints: ["web_search", "web_fetch", "web_image_search", "m365_try_office_sso", "m365_read_file"],
};

export default skill;
