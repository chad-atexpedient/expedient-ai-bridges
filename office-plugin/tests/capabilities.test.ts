import { describe, expect, it } from "vitest";
import { blockedToolResult, detectOfficeCapabilities, toolDefinitionsForCapabilities } from "../src/lib/capabilities";
import { OFFICE_TOOL_DEFINITIONS } from "../src/lib/tools";
import { buildAnthropicRequest, buildOpenAIRequest, completeChat } from "../src/providers/adapters";
import type { OfficeCapabilitySnapshot, ProviderSettings } from "../src/lib/types";

const base: ProviderSettings = {
  provider: "openai-compatible",
  apiKey: "test-key",
  baseUrl: "https://example.test/v1/",
  route: "openai-v1",
  modelMode: "auto",
  model: "GPT-5-Mini",
  autoModels: ["GPT-5-Mini", "GPT-5.4", "GPT-5.5"],
  temperature: null,
  maxTokens: null,
  useLocalProxy: true,
};

const messages = [{ id: "1", role: "user" as const, content: "Hello", timestamp: "now" }];
const context = { host: "word" as const, title: "Doc", selectionLabel: "Selection", text: "Selected text", metadata: {} };

const wordCapabilities: OfficeCapabilitySnapshot = {
  host: "word",
  platform: "PC",
  requirementSets: { WordApi_1_1: true },
  availableTools: ["web_search", "web_fetch", "web_image_search", "office_read_context", "ctrl_create_demo_showcase", "word_insert_text"],
  unavailableTools: {
    excel_add_worksheet: "Excel tool is only available in Excel.",
    excel_rename_worksheet: "Excel tool is only available in Excel.",
    excel_delete_worksheet: "Excel tool is only available in Excel.",
    excel_set_worksheet_visibility: "Excel tool is only available in Excel.",
    excel_clear_range: "Excel tool is only available in Excel.",
    excel_write_range: "Excel tool is only available in Excel.",
    excel_import_context_table: "Excel tool is only available in Excel.",
    excel_set_formula: "Excel tool is only available in Excel.",
    excel_clean_transform_range: "Excel tool is only available in Excel.",
    excel_combine_ranges: "Excel tool is only available in Excel.",
    excel_audit_formulas: "Excel tool is only available in Excel.",
    excel_map_formula_dependencies: "Excel tool is only available in Excel.",
    excel_create_chart: "Excel tool is only available in Excel.",
    excel_create_pivot_chart_report: "Excel tool is only available in Excel.",
    excel_create_pivottable: "Excel tool is only available in Excel.",
    powerpoint_create_slides: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_audit_deck: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_add_textbox: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_add_table: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_add_shape: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_create_diagram: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_set_slide_background: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_insert_image: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_generate_deck_file: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_delete_slide: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_clear_slide: "PowerPoint tool is only available in PowerPoint.",
    powerpoint_add_speaker_notes: "PowerPoint tool is only available in PowerPoint.",
    word_audit_document: "Word tool is only available in Word.",
    word_create_context_brief: "Word tool is only available in Word.",
  },
  notes: [],
};

describe("Office capability detection contract", () => {
  it("filters model tools to the current Office capability snapshot", () => {
    const definitions = toolDefinitionsForCapabilities(wordCapabilities);
    expect(definitions.map((tool) => tool.function.name)).toEqual(["web_search", "web_fetch", "web_image_search", "office_read_context", "ctrl_create_demo_showcase", "word_insert_text"]);

    const request = buildOpenAIRequest(base, messages, context, undefined, [], true, undefined, false, null, wordCapabilities);
    const body = JSON.parse(String(request.init.body));
    expect(body.tools.map((tool: any) => tool.function.name)).toEqual(["web_search", "web_fetch", "web_image_search", "office_read_context", "ctrl_create_demo_showcase", "word_insert_text"]);
    expect(body.messages[1].content).toContain("Office capability snapshot");
    expect(body.messages[1].content).toContain("Unavailable tools: excel_add_worksheet");
  });

  it("adds the same capability snapshot to Anthropic-compatible system context", () => {
    const request = buildAnthropicRequest({ ...base, provider: "anthropic-compatible", route: "anthropic-v1" }, messages, context, undefined, [], wordCapabilities);
    const body = JSON.parse(String(request.init.body));
    expect(body.system).toContain("Office capability snapshot");
    expect(body.system).toContain("Available tools: web_search, web_fetch, web_image_search, office_read_context, ctrl_create_demo_showcase, word_insert_text");
  });

  it("returns a clear blocked result for unavailable Office tools", () => {
    const result = blockedToolResult({ id: "chart", name: "excel_create_chart", arguments: {} }, wordCapabilities);
    expect(result?.ok).toBe(false);
    expect(result?.content).toContain("Excel tool is only available in Excel");
  });

  it("exposes workbook grounding tools in Excel and blocks them elsewhere", () => {
    const grounding = ["excel_get_workbook_overview", "excel_read_range", "excel_search_workbook"];
    const excelTools = OFFICE_TOOL_DEFINITIONS.map((tool) => tool.function.name);
    for (const name of grounding) expect(excelTools).toContain(name);

    const excelCapabilities: OfficeCapabilitySnapshot = {
      host: "excel",
      platform: "PC",
      requirementSets: { ExcelApi_1_1: true },
      availableTools: grounding,
      unavailableTools: {},
      notes: [],
    };
    expect(toolDefinitionsForCapabilities(excelCapabilities).map((tool) => tool.function.name)).toEqual(grounding);

    const inWord = detectOfficeCapabilities("word");
    for (const name of grounding) {
      expect(inWord.availableTools).not.toContain(name);
      expect(inWord.unavailableTools[name]).toBe("Excel tool is only available in Excel.");
    }
  });

  it("declares an explicit overwrite escape hatch on excel_write_range", () => {
    const writeTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "excel_write_range");
    const properties = (writeTool?.function.parameters as any)?.properties ?? {};
    expect(properties.overwrite?.type).toBe("boolean");
    expect(String(writeTool?.function.description)).toContain("refused");
    expect((writeTool?.function.parameters as any)?.required).not.toContain("overwrite");
  });

  it("can expose or hide native Excel PivotTable tooling based on runtime capability snapshots", () => {
    const supported: OfficeCapabilitySnapshot = {
      host: "excel",
      platform: "PC",
      requirementSets: { ExcelApi_1_1: true },
      availableTools: ["excel_create_pivottable", "excel_create_summary_table", "excel_generate_workbook_file"],
      unavailableTools: {},
      notes: [],
    };
    expect(toolDefinitionsForCapabilities(supported).map((tool) => tool.function.name)).toContain("excel_create_pivottable");

    const unsupported: OfficeCapabilitySnapshot = {
      ...supported,
      availableTools: ["excel_create_summary_table", "excel_generate_workbook_file"],
      unavailableTools: { excel_create_pivottable: "Excel PivotTable APIs are not available in this Office runtime. Use excel_create_summary_table or excel_generate_workbook_file as the fallback." },
    };
    expect(toolDefinitionsForCapabilities(unsupported).map((tool) => tool.function.name)).not.toContain("excel_create_pivottable");
    expect(blockedToolResult({ id: "pivot", name: "excel_create_pivottable", arguments: {} }, unsupported)?.content).toContain("excel_create_summary_table");
  });
  it("exposes the generated PowerPoint deck tool in PowerPoint capability snapshots", () => {
    const powerpointCapabilities: OfficeCapabilitySnapshot = {
      host: "powerpoint",
      platform: "PC",
      requirementSets: { PowerPointApi_1_1: true },
      availableTools: ["web_search", "web_fetch", "office_read_context", "powerpoint_generate_deck_file"],
      unavailableTools: {},
      notes: [],
    };

    const definitions = toolDefinitionsForCapabilities(powerpointCapabilities);
    expect(definitions.map((tool) => tool.function.name)).toContain("powerpoint_generate_deck_file");
  });

  it("hides live PowerPoint speaker notes when the runtime cannot prove notes support", () => {
    const previousOffice = globalThis.Office;
    const previousPowerPoint = globalThis.PowerPoint;
    try {
      (globalThis as any).Office = { context: { platform: "PC", requirements: { isSetSupported: () => true } } };
      (globalThis as any).PowerPoint = {};
      const capabilities = detectOfficeCapabilities("powerpoint");
      expect(capabilities.availableTools).toContain("powerpoint_generate_deck_file");
      expect(capabilities.availableTools).not.toContain("powerpoint_add_speaker_notes");
      expect(capabilities.unavailableTools.powerpoint_add_speaker_notes).toContain("Use powerpoint_generate_deck_file");
    } finally {
      (globalThis as any).Office = previousOffice;
      (globalThis as any).PowerPoint = previousPowerPoint;
    }
  });

  it("contains professional structure tools in the global tool schema", () => {
    const names = toolDefinitionsForCapabilities(null).map((tool) => tool.function.name);
    expect(names).toEqual(expect.arrayContaining([
      "excel_rename_worksheet",
      "excel_delete_worksheet",
      "excel_set_worksheet_visibility",
      "excel_clear_range",
      "excel_import_context_table",
      "excel_clean_transform_range",
      "excel_combine_ranges",
      "excel_audit_formulas",
      "excel_map_formula_dependencies",
      "excel_create_pivot_chart_report",
      "excel_create_pivottable",
      "word_insert_section_break",
      "word_audit_document",
      "word_create_context_brief",
      "word_set_header_footer",
      "powerpoint_set_slide_background",
      "powerpoint_create_diagram",
      "powerpoint_audit_deck",
      "powerpoint_generate_deck_file",
      "ctrl_create_demo_showcase",
      "excel_add_comment",
      "excel_set_named_range",
      "excel_protect_sheet",
      "excel_set_page_layout",
      "word_insert_hyperlink",
      "word_format_selection",
      "word_insert_content_control",
      "powerpoint_duplicate_slide",
      "powerpoint_add_hyperlink_textbox",
    ]));
  });

  it("does not force-nudge a native tool when capability detection marks it unavailable", async () => {
    let calls = 0;
    const noWebCapabilities: OfficeCapabilitySnapshot = {
      ...wordCapabilities,
      availableTools: ["office_read_context", "word_insert_text"],
      unavailableTools: { ...wordCapabilities.unavailableTools, web_search: "Web search is disabled by policy.", web_fetch: "Web fetch is disabled by policy.", web_image_search: "Web image search is disabled by policy." },
    };

    const result = await completeChat({
      settings: base,
      messages: [{ ...messages[0], content: "Look up Colorado gas prices over the last 12 months" }],
      capabilities: noWebCapabilities,
      toolsEnabled: true,
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "native search result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        expect(body.tools.map((tool: any) => tool.function.name)).not.toContain("web_search");
        expect(body.tool_choice).toBe("auto");
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "web search is unavailable here" } }] }), { status: 200 });
      },
    });

    expect(calls).toBe(1);
    expect(result.text).toBe("web search is unavailable here");
  });
});
