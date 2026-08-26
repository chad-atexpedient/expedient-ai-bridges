import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "word-editor",
  name: "Word Editor",
  description: "Edits and reviews Word documents with audit tables, context briefs, headings, text, tables, comments, images, links, headers/footers, and generated .docx fallbacks.",
  appliesTo: ["word"],
  priority: 100,
  instructions: [
    "Treat document requests as Word editing or generated document tasks.",
    "For document QA, accessibility, readability, structure, or professional polish reviews, use word_audit_document to insert a real review table before giving generic advice.",
    "When the user asks to turn Excel, PowerPoint, Microsoft 365, uploaded, web, or chat context into a Word memo, brief, pre-read, or summary, use word_create_context_brief so the current document gets structured headings, summary, key points, and source table where available.",
    "Use headings, tables, comments, formatting, content controls/template fields, and generated document files for professional structure rather than plain paragraphs when the user asks for a document artifact.",
    "Use uploaded or web image assets as real images when available.",
    "Use word_generate_document_file when a standalone report or template-like artifact is more reliable than live mutation. When the user attaches a DOCX template, pass templateAssetId or templateAssetName so the package shell, styles, theme, custom XML, headers/footers, and reusable media can be preserved while the generated body is replaced.",
  ],
  toolHints: ["word_audit_document", "word_create_context_brief", "word_insert_heading", "word_insert_table", "word_insert_image", "word_insert_comment", "word_insert_content_control", "word_generate_document_file"],
};

export default skill;
