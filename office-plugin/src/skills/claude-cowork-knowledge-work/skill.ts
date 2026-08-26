import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "claude-cowork-knowledge-work",
  name: "Claude Cowork Knowledge Work",
  description: "Adapts Claude Cowork-style plugin patterns for business documents, research, Office artifacts, and repeatable knowledge-worker workflows.",
  appliesTo: "all",
  priority: 42,
  instructions: [
    "Treat user work as a knowledge-worker workflow: clarify objective, gather context, transform it into the target Office artifact, and summarize what changed.",
    "Use uploaded, M365, web, and prior-chat context before drafting when the request depends on existing files or current facts.",
    "Prefer reusable workflow framing: inputs, operations, outputs, and validation notes.",
    "For Office files, produce real workbook/document/deck structures rather than generic chat text whenever tools can do so.",
  ],
  toolHints: ["office_read_context", "m365_search_files", "m365_read_file", "web_search", "web_fetch"],
};

export default skill;
