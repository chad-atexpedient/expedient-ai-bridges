import type { SkillDefinition } from "../types";

const skill: SkillDefinition = {
  id: "excel-operator",
  name: "Excel Operator",
  description: "Directly edits Excel workbooks with native worksheet, range, context import, cleanup/transform, append/merge, table, formula, formula-audit/dependency-map, chart, PivotTable-style report, formatting, and image tools.",
  appliesTo: ["excel"],
  priority: 100,
  instructions: [
    "Treat spreadsheet requests as direct workbook operations whenever Excel tools are available.",
    "Use local A1 ranges such as A1:C12 when sheetName is supplied; do not include the sheet prefix in address arguments.",
    "For current facts, search/fetch first, then write structured rows and create charts, native PivotTables when available, summary-table fallbacks, or structured tables as real Excel objects.",
    "When the user asks to use Word, PowerPoint, Microsoft 365, uploaded, web, or chat context in Excel, read the relevant context first, then use excel_import_context_table to create a structured worksheet table with source provenance.",
    "For data cleanup, CSV/web imports, split-column, whitespace cleanup, blank-row removal, or dedupe requests, use excel_clean_transform_range to create a cleaned output table instead of giving manual Power Query/Text to Columns instructions.",
    "For append, combine, consolidate, merge, or lookup-join requests across two ranges, use excel_combine_ranges to create a combined output table instead of manual copy/paste or Power Query instructions.",
    "For pivot-chart, grouped chart, visual summary, or executive analyst view requests, use excel_create_pivot_chart_report to create a grouped table and chart in one report sheet when a full native PivotTable is unnecessary or unavailable.",
    "For workbook audit, risk review, formula review, or finance-model QA requests, use excel_audit_formulas to create a Formula Audit sheet and excel_map_formula_dependencies when lineage, precedents, dependents, or model flow matter.",
    "If a live mutation fails twice because of runtime limits, use excel_generate_workbook_file rather than giving manual instructions.",
  ],
  toolHints: ["excel_write_range", "excel_import_context_table", "excel_set_formula", "excel_clean_transform_range", "excel_combine_ranges", "excel_audit_formulas", "excel_map_formula_dependencies", "excel_create_chart", "excel_create_pivot_chart_report", "excel_create_pivottable", "excel_create_summary_table", "excel_create_table", "excel_format_range", "excel_set_page_layout", "excel_generate_workbook_file"],
};

export default skill;
