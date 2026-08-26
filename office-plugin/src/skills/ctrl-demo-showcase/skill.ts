import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "ctrl-demo-showcase",
  name: "CTRL Demo Showcase",
  description: "Creates prepared CTRL demonstrations live in the active Office file by default, with downloadable Excel, Word, and PowerPoint artifacts available as an explicit fallback or full-suite mode.",
  appliesTo: "all",
  priority: 80,
  instructions: [
    "Simple demo/showcase/tour requests should render as an in-chat add-in demo card. Use ctrl_create_demo_showcase only when the user clicks or asks for a live Office sample, downloadable artifact, or full suite.",
    "Use surface=current and mode=live for the active Office app when the user explicitly wants a small live sample in the current workbook, document, or deck.",
    "Use surface=all or mode=artifact when the user asks for a full suite, downloadable artifact, preview-safe output, or when live Office APIs are unavailable.",
    "After creating the demo, briefly name whether it was live or artifact-based and the major features shown.",
  ],
  toolHints: ["ctrl_create_demo_showcase", "excel_generate_workbook_file", "word_generate_document_file", "powerpoint_generate_deck_file"],
};

export default skill;
