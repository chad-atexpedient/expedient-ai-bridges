import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "office-generation",
  name: "Generated Office Artifacts",
  description: "Creates downloadable .xlsx, .docx, and .pptx artifacts when live Office APIs are too limited or a standalone file is better.",
  appliesTo: "all",
  priority: 50,
  instructions: [
    "Prefer live Office mutation for simple edits in the active file, but use generated Office files for richer standalone artifacts or runtime API gaps.",
    "Generated artifacts must contain real Office objects where supported: tables, charts, notes, images, hyperlinks, comments, and layout structure.",
    "Do not return manual build instructions when a generated file tool can create the artifact.",
  ],
  toolHints: ["excel_generate_workbook_file", "word_generate_document_file", "powerpoint_generate_deck_file"],
};

export default skill;
