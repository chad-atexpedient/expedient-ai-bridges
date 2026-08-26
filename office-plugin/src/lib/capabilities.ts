import type { OfficeCapabilitySnapshot, OfficeHost, ToolCallRequest, ToolCallResult } from "./types";
import { OFFICE_TOOL_DEFINITIONS } from "./tools";

const WEB_TOOLS = new Set(["web_search", "web_fetch", "web_image_search", "web_image_import"]);
const M365_TOOLS = new Set(["m365_try_office_sso", "m365_auth_status", "m365_search_files", "m365_read_file"]);
const SHARED_OFFICE_TOOLS = new Set(["office_read_context", "ctrl_create_demo_showcase", "word_generate_document_file", "excel_generate_workbook_file"]);
const EXCEL_TOOLS = new Set(["excel_get_workbook_overview", "excel_read_range", "excel_search_workbook", "excel_add_worksheet", "excel_rename_worksheet", "excel_delete_worksheet", "excel_set_worksheet_visibility", "excel_clear_range", "excel_write_range", "excel_import_context_table", "excel_set_formula", "excel_clean_transform_range", "excel_combine_ranges", "excel_audit_formulas", "excel_map_formula_dependencies", "excel_create_chart", "excel_create_summary_table", "excel_create_pivot_chart_report", "excel_create_pivottable", "excel_create_table", "excel_format_range", "excel_sort_range", "excel_filter_range", "excel_freeze_panes", "excel_add_data_validation", "excel_apply_conditional_format", "excel_insert_image", "excel_add_comment", "excel_set_named_range", "excel_protect_sheet", "excel_set_page_layout"]);
const WORD_TOOLS = new Set(["word_audit_document", "word_create_context_brief", "word_insert_text", "word_insert_heading", "word_insert_table", "word_apply_style", "word_insert_page_break", "word_insert_section_break", "word_set_header_footer", "word_find_replace", "word_insert_comment", "word_insert_image", "word_insert_hyperlink", "word_format_selection", "word_insert_content_control"]);
const POWERPOINT_TOOLS = new Set(["powerpoint_audit_deck", "powerpoint_create_slides", "powerpoint_add_textbox", "powerpoint_add_table", "powerpoint_add_shape", "powerpoint_create_diagram", "powerpoint_arrange_shapes", "powerpoint_group_shapes", "powerpoint_delete_slide", "powerpoint_clear_slide", "powerpoint_set_slide_background", "powerpoint_add_speaker_notes", "powerpoint_insert_image", "powerpoint_duplicate_slide", "powerpoint_add_hyperlink_textbox", "powerpoint_generate_deck_file"]);

function requirementSupported(name: string, version: string) {
  try {
    return Boolean(globalThis.Office?.context?.requirements?.isSetSupported?.(name, version));
  } catch {
    return false;
  }
}

function platformName() {
  return String(globalThis.Office?.context?.platform || globalThis.Office?.PlatformType?.OfficeOnline || "unknown");
}

function hasExcelChartApi() {
  return Boolean(globalThis.Excel?.ChartType && globalThis.Excel?.ChartSeriesBy);
}

function hasExcelPivotTableApi() {
  return Boolean(globalThis.Excel?.PivotLayoutType || globalThis.Excel?.AggregationFunction || globalThis.Excel?.PivotAxis);
}

function hasRequirementOrApi(requirementName: string, version: string, apiObject: unknown) {
  if (!apiObject) return false;
  const requirements = globalThis.Office?.context?.requirements;
  if (!requirements?.isSetSupported) return true;
  return requirementSupported(requirementName, version);
}

function toolUnavailableReason(host: OfficeHost, toolName: string) {
  if (WEB_TOOLS.has(toolName)) return null;
  if (M365_TOOLS.has(toolName)) return null;
  if (SHARED_OFFICE_TOOLS.has(toolName)) return toolName === "office_read_context" && host === "unknown" ? "Office context is not connected." : null;
  if (EXCEL_TOOLS.has(toolName)) {
    if (host !== "excel") return "Excel tool is only available in Excel.";
    if (!globalThis.Excel) return "Excel JavaScript APIs are not loaded.";
    if (!hasRequirementOrApi("ExcelApi", "1.1", globalThis.Excel)) return "ExcelApi 1.1 is not supported in this Office runtime.";
    if (toolName === "excel_create_chart" && !hasExcelChartApi()) return "Excel chart APIs are not available in this Office runtime.";
    if (toolName === "excel_create_pivottable" && !hasExcelPivotTableApi()) return "Excel PivotTable APIs are not available in this Office runtime. Use excel_create_summary_table or excel_generate_workbook_file as the fallback.";
    if (toolName === "excel_insert_image" && !globalThis.Excel) return "Excel image APIs are not available in this Office runtime.";
    return null;
  }
  if (WORD_TOOLS.has(toolName)) {
    if (host !== "word") return "Word tool is only available in Word.";
    if (!globalThis.Word) return "Word JavaScript APIs are not loaded.";
    if (!hasRequirementOrApi("WordApi", "1.1", globalThis.Word)) return "WordApi 1.1 is not supported in this Office runtime.";
    return null;
  }
  if (POWERPOINT_TOOLS.has(toolName)) {
    if (toolName === "powerpoint_generate_deck_file") return null;
    if (toolName === "powerpoint_add_speaker_notes") return "PowerPoint speaker notes APIs are not reliably exposed to task-pane add-ins in this runtime. Use powerpoint_generate_deck_file for guaranteed speaker notes.";
    if (host !== "powerpoint") return "PowerPoint tool is only available in PowerPoint.";
    if (!globalThis.PowerPoint) return "PowerPoint JavaScript APIs are not loaded.";
    if (!hasRequirementOrApi("PowerPointApi", "1.1", globalThis.PowerPoint)) return "PowerPointApi 1.1 is not supported in this Office runtime.";
    return null;
  }
  return "Unknown tool.";
}

export function detectOfficeCapabilities(host: OfficeHost): OfficeCapabilitySnapshot {
  const unavailableTools: Record<string, string> = {};
  const availableTools: string[] = [];
  for (const tool of OFFICE_TOOL_DEFINITIONS) {
    const reason = toolUnavailableReason(host, tool.function.name);
    if (reason) unavailableTools[tool.function.name] = reason;
    else availableTools.push(tool.function.name);
  }

  const requirementSets = {
    ExcelApi_1_1: requirementSupported("ExcelApi", "1.1"),
    ExcelApi_1_8: requirementSupported("ExcelApi", "1.8"),
    WordApi_1_1: requirementSupported("WordApi", "1.1"),
    PowerPointApi_1_1: requirementSupported("PowerPointApi", "1.1"),
  };

  const notes: string[] = [];
  if (host === "excel" && !availableTools.includes("excel_create_chart")) notes.push(unavailableTools.excel_create_chart || "Excel chart tool unavailable.");
  if (host === "excel" && !availableTools.includes("excel_create_pivottable")) notes.push(unavailableTools.excel_create_pivottable || "Excel PivotTable tool unavailable; use summary table fallback.");
  if (host === "word" && !availableTools.includes("word_insert_text")) notes.push(unavailableTools.word_insert_text || "Word insert tool unavailable.");
  if (host === "powerpoint" && !availableTools.includes("powerpoint_create_slides")) notes.push(unavailableTools.powerpoint_create_slides || "PowerPoint slide tool unavailable.");
  if (host === "unknown") notes.push("Running in preview/browser mode; Office document mutation tools are unavailable.");

  return {
    host,
    platform: platformName(),
    requirementSets,
    availableTools,
    unavailableTools,
    notes,
  };
}

export function toolDefinitionsForCapabilities(capabilities?: OfficeCapabilitySnapshot | null) {
  if (!capabilities) return OFFICE_TOOL_DEFINITIONS;
  const allowed = new Set(capabilities.availableTools);
  return OFFICE_TOOL_DEFINITIONS.filter((tool) => allowed.has(tool.function.name));
}

export function isToolAvailable(capabilities: OfficeCapabilitySnapshot | null | undefined, toolName: string) {
  if (!capabilities) return true;
  return capabilities.availableTools.includes(toolName);
}

export function blockedToolResult(call: ToolCallRequest, capabilities: OfficeCapabilitySnapshot): ToolCallResult | null {
  if (isToolAvailable(capabilities, call.name)) return null;
  return {
    id: call.id,
    name: call.name,
    ok: false,
    content: capabilities.unavailableTools[call.name] || `Tool ${call.name} is not available in this Office host.`,
  };
}

export function capabilityPrompt(capabilities?: OfficeCapabilitySnapshot | null) {
  if (!capabilities) return null;
  return [
    `Office capability snapshot:`,
    `Host: ${capabilities.host}`,
    `Platform: ${capabilities.platform}`,
    `Available tools: ${capabilities.availableTools.join(", ") || "none"}`,
    `Unavailable tools: ${Object.entries(capabilities.unavailableTools).map(([tool, reason]) => `${tool} (${reason})`).join("; ") || "none"}`,
    capabilities.notes.length ? `Notes: ${capabilities.notes.join(" ")}` : "",
  ].filter(Boolean).join("\n");
}


