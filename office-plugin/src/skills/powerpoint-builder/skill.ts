import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "powerpoint-builder",
  name: "PowerPoint Builder",
  description: "Builds, edits, and reviews decks with real audit slides, slides, text boxes, tables, shapes, images, backgrounds, notes, and generated/template-preserving .pptx fallbacks.",
  appliesTo: ["powerpoint"],
  priority: 100,
  instructions: [
    "Treat deck requests as slide operations in PowerPoint, not as generic prose drafting.",
    "For deck QA, design review, accessibility/design-readiness, executive polish, or slide-story review, use powerpoint_audit_deck to add a concrete audit slide before giving generic presentation advice.",
    "For process, cycle, roadmap, timeline, workflow, or simple framework diagrams, use powerpoint_create_diagram to create live editable slide objects instead of giving SmartArt instructions.",
    "Use slideNumber for human-facing slide references, keep layout coordinates positive and simple, and use arrange/group tools to align, distribute, order, group, or ungroup objects after creation.",
    "Use uploaded or web image assets as real images; never create image placeholders when an image URL or generated deck can embed the media.",
    "When attached PPTX context lists layout names or types, choose layoutName/layoutType per generated slide instead of forcing every slide through one generic layout.",
    "When an uploaded PPTX should act as the brand/template source, pass its attachment id or filename as templateAssetId/templateAssetName to powerpoint_generate_deck_file so theme, masters, layouts, relationships, fonts, reusable media, text placeholders, and image/table/chart placeholder placements are preserved or filled where available.",
    "For generated deck charts, provide editable chart options such as chart type, axis titles, number format, labels, colors, legend position, horizontal bars, and gridline visibility when they improve executive polish. Use categories plus series for multi-series comparisons across years, regions, products, scenarios, or groups.",
    "When live PowerPoint APIs cannot create the requested real object, use powerpoint_generate_deck_file instead of manual instructions.",
  ],
  toolHints: ["powerpoint_audit_deck", "powerpoint_create_slides", "powerpoint_add_textbox", "powerpoint_add_table", "powerpoint_add_shape", "powerpoint_create_diagram", "powerpoint_arrange_shapes", "powerpoint_group_shapes", "powerpoint_insert_image", "powerpoint_generate_deck_file"],
};

export default skill;
