import { describe, expect, it } from "vitest";
import { buildAnthropicRequest, buildOpenAIRequest, completeChat, selectModel } from "../src/providers/adapters";
import { OFFICE_TOOL_DEFINITIONS } from "../src/lib/tools";
import type { ProviderSettings } from "../src/lib/types";

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

describe("provider adapters", () => {
  it("exposes native Excel chart creation as a model tool", () => {
    const chartTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "excel_create_chart");
    const required = "required" in chartTool!.function.parameters ? chartTool!.function.parameters.required : [];
    expect(required).toContain("sourceAddress");
    expect(required).toContain("chartType");
    expect(Object.keys(chartTool!.function.parameters.properties)).toContain("startCell");
    expect(Object.keys(chartTool!.function.parameters.properties)).toContain("endCell");
  });

  it("exposes rich generated Excel workbook chart options", () => {
    const workbookTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "excel_generate_workbook_file") as any;
    expect(Object.keys(workbookTool.function.parameters.properties)).toEqual(expect.arrayContaining(["title", "author", "subject", "properties", "namedRanges"]));
    expect(Object.keys(workbookTool.function.parameters.properties.properties.properties)).toEqual(expect.arrayContaining(["title", "subject", "creator", "author", "keywords", "description", "category", "company", "manager"]));
    expect(Object.keys(workbookTool.function.parameters.properties.namedRanges.items.properties)).toEqual(expect.arrayContaining(["name", "sheetName", "sheetIndex", "address", "range", "ref", "reference"]));
    const sheetProps = workbookTool.function.parameters.properties.sheets.items.properties;
    expect(Object.keys(sheetProps)).toEqual(expect.arrayContaining(["columns", "rowHeights", "rowsMeta", "columnWidths", "merges", "mergeCells", "autoFilter", "filter", "namedRanges", "freeze", "freezePanes", "freezeRows", "freezeColumns", "zoomScale", "orientation", "paperSize", "margins", "printArea", "repeatRows", "repeatColumns", "fitToPagesWide", "fitToPagesTall", "showGridlines", "showHeadings", "centerHorizontally", "centerVertically", "blackAndWhite", "draftMode", "protect", "password", "protection", "allowFormatCells", "allowSort", "allowAutoFilter", "comments", "notes", "links", "hyperlinks"]));
    expect(Object.keys(sheetProps.namedRanges.items.properties)).toEqual(expect.arrayContaining(["name", "address", "range", "ref", "reference", "scope"]));
    expect(Object.keys(sheetProps.protection.properties)).toEqual(expect.arrayContaining(["protect", "protected", "password", "allowFormatCells", "allowSort", "allowAutoFilter"]));
    expect(Object.keys(sheetProps.comments.items.properties)).toEqual(expect.arrayContaining(["address", "cell", "text", "author", "visible"]));
    expect(Object.keys(sheetProps.links.items.properties)).toEqual(expect.arrayContaining(["address", "cell", "url", "target", "location", "text", "label", "display", "tooltip", "screenTip"]));
    const chartSchema = workbookTool.function.parameters.properties.sheets.items.properties.charts.items;
    const chartProps = chartSchema.properties;
    expect(chartProps.chartType.enum).toEqual(expect.arrayContaining(["bar", "line", "pie", "area", "doughnut", "scatter", "combo"]));
    expect(Object.keys(chartProps)).toEqual(expect.arrayContaining([
      "categories",
      "series",
      "lineSeries",
      "lineSeriesStartIndex",
      "points",
      "categoryAxisTitle",
      "valueAxisTitle",
      "valueFormat",
      "xValueFormat",
      "yValueFormat",
      "dataLabels",
      "scatterStyle",
      "colors",
      "legendPosition",
      "barDirection",
      "showGridLines",
      "holeSize",
    ]));
  });

  it("exposes rich generated PowerPoint image and table options", () => {
    const deckTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "powerpoint_generate_deck_file") as any;
    const deckProps = deckTool.function.parameters.properties;
    expect(Object.keys(deckProps)).toEqual(expect.arrayContaining(["footer", "footerText", "dateText", "date", "showDate", "showFooter", "showSlideNumber", "showSlideNumbers", "slideNumberFormat", "slideNumberText", "confidentialityLabel", "footerColor", "footerFontSize"]));
    const slideProps = deckTool.function.parameters.properties.slides.items.properties;
    expect(Object.keys(slideProps)).toEqual(expect.arrayContaining(["footer", "footerText", "dateText", "date", "showDate", "showFooter", "showSlideNumber", "showSlideNumbers", "slideNumberFormat", "slideNumberText", "confidentialityLabel", "footerColor", "footerFontSize"]));
    const imageProps = slideProps.images.items.properties;
    expect(Object.keys(imageProps)).toEqual(expect.arrayContaining(["pixelWidth", "pixelHeight", "fit", "sizing", "crop", "cropLeft", "cropRight", "cropTop", "cropBottom"]));
    expect(imageProps.fit.enum).toEqual(expect.arrayContaining(["stretch", "fit", "contain", "fill", "cover"]));
    expect(Object.keys(imageProps.crop.properties)).toEqual(expect.arrayContaining(["left", "right", "top", "bottom"]));
    const tableProps = slideProps.tables.items.properties;
    expect(Object.keys(tableProps)).toEqual(expect.arrayContaining(["headerFillColor", "headerTextColor", "bodyFillColor", "bandFillColor", "alternateRowFillColor", "textColor", "borderColor", "borderWidth", "fontSize", "headerFontSize", "headerBold", "bold", "align", "headerAlign", "verticalAlign", "firstRow", "bandRows", "columnWidths", "rowHeights"]));
    expect(tableProps.align.enum).toEqual(expect.arrayContaining(["left", "center", "right", "justify"]));
  });

  it("exposes professional Office authoring tools across hosts", () => {
    const names = OFFICE_TOOL_DEFINITIONS.map((tool) => tool.function.name);
    const pageLayoutTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "excel_set_page_layout");
    expect(Object.keys(pageLayoutTool!.function.parameters.properties)).toEqual(expect.arrayContaining(["printArea", "repeatRows", "repeatColumns", "fitToPagesWide", "fitToPagesTall", "showGridlines", "showHeadings", "centerHorizontally", "centerVertically"]));
    const contentControlTool = OFFICE_TOOL_DEFINITIONS.find((tool) => tool.function.name === "word_insert_content_control");
    expect(Object.keys(contentControlTool!.function.parameters.properties)).toEqual(expect.arrayContaining(["title", "tag", "placeholderText", "type", "options", "cannotDelete", "cannotEdit"]));
    expect(names).toEqual(expect.arrayContaining([
      "excel_create_pivottable",
      "ctrl_create_demo_showcase",
      "excel_create_table",
      "excel_create_summary_table",
      "web_image_search",
      "web_image_import",
      "excel_format_range",
      "excel_sort_range",
      "excel_filter_range",
      "excel_freeze_panes",
      "excel_add_data_validation",
      "excel_apply_conditional_format",
      "excel_add_comment",
      "excel_set_named_range",
      "excel_protect_sheet",
      "excel_set_page_layout",
      "word_insert_heading",
      "word_insert_table",
      "word_apply_style",
      "word_insert_page_break",
      "word_find_replace",
      "word_insert_comment",
      "word_insert_hyperlink",
      "word_format_selection",
      "word_insert_content_control",
      "powerpoint_add_textbox",
      "powerpoint_add_table",
      "powerpoint_add_shape",
      "powerpoint_arrange_shapes",
      "powerpoint_group_shapes",
      "powerpoint_delete_slide",
      "powerpoint_clear_slide",
      "powerpoint_duplicate_slide",
      "powerpoint_add_hyperlink_textbox",
      "powerpoint_add_speaker_notes",
      "word_generate_document_file",
      "excel_generate_workbook_file",
      "powerpoint_insert_image",
      "powerpoint_generate_deck_file",
    ]));
  });

  it("builds proxied OpenAI-compatible requests and omits gateway-owned tuning defaults", () => {
    const request = buildOpenAIRequest(base, messages, context);
    expect(request.url).toBe("/api/proxy?target=https%3A%2F%2Fexample.test%2Fv1%2Fchat%2Fcompletions");
    expect((request.init.headers as Record<string, string>)["x-provider-authorization"]).toBe("Bearer test-key");
    const body = JSON.parse(String(request.init.body));
    expect(body.model).toBe("GPT-5-Mini");
    expect(body.temperature).toBeUndefined();
    expect(body.max_tokens).toBeUndefined();
    expect(body.messages[1].content).toContain("Selected text");
  });

  it("includes temperature and max tokens only when explicitly set", () => {
    const request = buildOpenAIRequest({ ...base, temperature: 1, maxTokens: 2000 }, messages, context);
    const body = JSON.parse(String(request.init.body));
    expect(body.temperature).toBe(1);
    expect(body.max_tokens).toBe(2000);
    expect(body.maxTokens).toBeUndefined();
  });

  it("keeps Anthropic-compatible tuning fields gateway-safe", () => {
    const request = buildAnthropicRequest({ ...base, provider: "anthropic-compatible", route: "anthropic-v1", temperature: null, maxTokens: null }, messages, context);
    const body = JSON.parse(String(request.init.body));
    expect(body.temperature).toBeUndefined();
    expect(body.max_tokens).toBeUndefined();
    expect(body.maxTokens).toBeUndefined();
  });
  it("includes Office context warnings without requiring context text", () => {
    const request = buildOpenAIRequest(base, messages, null, "Could not read the current Office context");
    const body = JSON.parse(String(request.init.body));
    expect(body.messages[1].content).toContain("Office context warning");
  });

  it("can build direct OpenAI-compatible requests when proxying is disabled", () => {
    const request = buildOpenAIRequest({ ...base, useLocalProxy: false }, messages, context);
    expect(request.url).toBe("https://example.test/v1/chat/completions");
  });

  it("supports custom full OpenWebUI chat completion URLs", () => {
    const request = buildOpenAIRequest({ ...base, baseUrl: "https://internal-beta.expedient.cloud/api/chat/completions", route: "custom" }, messages, context);
    expect(request.url).toBe("/api/proxy?target=https%3A%2F%2Finternal-beta.expedient.cloud%2Fapi%2Fchat%2Fcompletions");
  });

  it("auto-routes between the three recommended models", () => {
    expect(selectModel(base, messages, null)).toBe("GPT-5-Mini");
    expect(selectModel(base, [{ ...messages[0], content: "Summarize this sheet" }], { ...context, text: "x".repeat(1500) })).toBe("GPT-5.4");
    expect(selectModel(base, [{ ...messages[0], content: "Build a complex financial model and audit formulas" }], context)).toBe("GPT-5.5");
  });

  it("manual model selection overrides auto-routing", () => {
    expect(selectModel({ ...base, modelMode: "manual", model: "custom-model" }, [{ ...messages[0], content: "Build a complex workflow" }], context)).toBe("custom-model");
  });

  it("builds proxied Anthropic-compatible messages requests", () => {
    const request = buildAnthropicRequest({ ...base, provider: "anthropic-compatible", route: "anthropic-v1" }, messages, context);
    expect(request.url).toBe("/api/proxy?target=https%3A%2F%2Fexample.test%2Fv1%2Fmessages");
    expect((request.init.headers as Record<string, string>)["x-provider-api-key"]).toBe("test-key");
    const body = JSON.parse(String(request.init.body));
    expect(body.system).toContain("Selected text");
    expect(body.messages[0].role).toBe("user");
  });

  it("parses OpenAI-compatible responses", async () => {
    const result = await completeChat({ settings: base, messages }, {
      fetch: async () => new Response(JSON.stringify({ choices: [{ message: { content: "ok" } }] }), { status: 200 }),
    });
    expect(result.text).toBe("ok");
  });

  it("runs an OpenAI-compatible tool loop", async () => {
    let calls = 0;
    const statuses: string[] = [];
    const result = await completeChat({
      settings: base,
      messages,
      toolsEnabled: true,
      onStatus: (status) => statuses.push(status),
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "tool result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (calls === 1) {
          expect(body.tools.length).toBeGreaterThan(0);
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [{ id: "call-1", type: "function", function: { name: "web_search", arguments: "{\"query\":\"Colorado gas prices\"}" } }] } }] }), { status: 200 });
        }
        expect(body.messages.some((message: any) => message.role === "tool" && message.content === "tool result")).toBe(true);
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "done" } }] }), { status: 200 });
      },
    });
    expect(result.text).toBe("done");
    expect(result.toolResults?.[0]?.name).toBe("web_search");
    expect(statuses.some((status) => status.includes("Contacting model"))).toBe(true);
    expect(statuses.some((status) => status.includes("Running web_search"))).toBe(true);
  });

  it("nudges native web_search tool choice when current web data is requested", async () => {
    let calls = 0;
    const result = await completeChat({
      settings: base,
      messages: [{ ...messages[0], content: "Look up Colorado gas prices over the last 12 months" }],
      toolsEnabled: true,
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "native search result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (calls === 1) {
          expect(body.tool_choice).toBe("auto");
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "I should search." } }] }), { status: 200 });
        }
        if (calls === 2) {
          expect(body.tool_choice).toEqual({ type: "function", function: { name: "web_search" } });
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [{ id: "forced-search", type: "function", function: { name: "web_search", arguments: "{\"query\":\"Colorado gas prices last 12 months\"}" } }] } }] }), { status: 200 });
        }
        expect(body.messages.some((message: any) => message.role === "tool" && message.content === "native search result")).toBe(true);
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "searched" } }] }), { status: 200 });
      },
    });
    expect(result.text).toBe("searched");
    expect(result.toolResults?.[0]?.name).toBe("web_search");
  });

  it("nudges native web_image_search tool choice for visual asset requests", async () => {
    let calls = 0;
    const result = await completeChat({
      settings: base,
      messages: [{ ...messages[0], content: "Create a dolphin conservation deck with real dolphin images" }],
      toolsEnabled: true,
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "image search result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (calls === 1) {
          expect(body.tool_choice).toBe("auto");
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "I need image candidates." } }] }), { status: 200 });
        }
        if (calls === 2) {
          expect(body.tool_choice).toEqual({ type: "function", function: { name: "web_image_search" } });
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [{ id: "forced-image-search", type: "function", function: { name: "web_image_search", arguments: "{\"query\":\"dolphin conservation photos\"}" } }] } }] }), { status: 200 });
        }
        expect(body.messages.some((message: any) => message.role === "tool" && message.content === "image search result")).toBe(true);
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "images found" } }] }), { status: 200 });
      },
    });
    expect(result.text).toBe("images found");
    expect(result.toolResults?.[0]?.name).toBe("web_image_search");
  });

  it("forces a final answer when a model keeps calling tools", async () => {
    let calls = 0;
    const result = await completeChat({
      settings: base,
      messages,
      toolsEnabled: true,
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "tool result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (body.tools) {
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [{ id: `call-${calls}`, type: "function", function: { name: "web_search", arguments: "{\"query\":\"again\"}" } }] } }] }), { status: 200 });
        }
        expect(body.messages.at(-1).content).toContain("Stop calling tools now");
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "final summary" } }] }), { status: 200 });
      },
    });
    expect(result.text).toBe("final summary");
    expect(result.toolResults?.length).toBe(12);
  });

  it("uses generated artifact markdown links when the model gives no final summary", async () => {
    let calls = 0;
    const result = await completeChat({
      settings: base,
      messages: [{ ...messages[0], content: "Create a PowerPoint demo" }],
      toolsEnabled: true,
      toolExecutor: async (call) => ({
        id: call.id,
        name: call.name,
        ok: true,
        content: JSON.stringify({ ok: true, downloadUrl: "/api/generated/files/demo.pptx", markdownLink: "[Download the PowerPoint artifact](/api/generated/files/demo.pptx)" }),
      }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (calls === 1) {
          return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [{ id: "demo-call", type: "function", function: { name: "ctrl_create_demo_showcase", arguments: "{\"surface\":\"powerpoint\"}" } }] } }] }), { status: 200 });
        }
        if (body.tools) return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null, tool_calls: [] } }] }), { status: 200 });
        return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: null } }] }), { status: 200 });
      },
    });
    expect(result.text).toContain("[Download the PowerPoint artifact](/api/generated/files/demo.pptx)");
  });


  it("adds host runtime context to the latest OpenAI-compatible user prompt", () => {
    const request = buildOpenAIRequest(base, messages, { ...context, host: "excel", title: "Budget.xlsx", selectionLabel: "Sheet1!A1:C5" });
    const body = JSON.parse(String(request.init.body));
    const latestUser = body.messages.at(-1).content;
    expect(latestUser).toContain("Current Office surface: Excel");
    expect(latestUser).toContain("Current file/title: Budget.xlsx");
    expect(latestUser).toContain("Required behavior: understand the user's request as a task in Excel");
  });

  it("adds host runtime context to Anthropic-compatible user prompts", () => {
    const request = buildAnthropicRequest({ ...base, provider: "anthropic-compatible", route: "anthropic-v1" }, messages, { ...context, host: "excel", title: "Budget.xlsx", selectionLabel: "Sheet1!A1:C5" });
    const body = JSON.parse(String(request.init.body));
    expect(body.system).toContain("Selected text");
    expect(body.messages[0].content).toContain("Current Office surface: Excel");
    expect(body.messages[0].content).toContain("Available native tools for this surface");
  });

  it("adds Anthropic-compatible native tool schemas when tools are enabled", () => {
    const request = buildAnthropicRequest({ ...base, provider: "anthropic-compatible", route: "anthropic-v1" }, messages, context, undefined, [], undefined, true);
    const body = JSON.parse(String(request.init.body));
    expect(body.tools.length).toBeGreaterThan(0);
    const searchTool = body.tools.find((tool: any) => tool.name === "web_search");
    expect(searchTool.description).toContain("Search the web");
    expect(searchTool.input_schema.properties.query.type).toBe("string");
  });

  it("runs an Anthropic-compatible tool loop", async () => {
    let calls = 0;
    const statuses: string[] = [];
    const result = await completeChat({
      settings: { ...base, provider: "anthropic-compatible", route: "anthropic-v1" },
      messages,
      toolsEnabled: true,
      onStatus: (status) => statuses.push(status),
      toolExecutor: async (call) => ({ id: call.id, name: call.name, ok: true, content: "anthropic tool result" }),
    }, {
      fetch: async (_url, init) => {
        calls += 1;
        const body = JSON.parse(String(init?.body));
        if (calls === 1) {
          expect(body.tools.length).toBeGreaterThan(0);
          return new Response(JSON.stringify({
            content: [
              { type: "text", text: "I will search." },
              { type: "tool_use", id: "toolu_1", name: "web_search", input: { query: "Colorado gas prices" } },
            ],
          }), { status: 200 });
        }
        expect(body.messages.at(-1).role).toBe("user");
        expect(body.messages.at(-1).content[0]).toEqual({ type: "tool_result", tool_use_id: "toolu_1", content: "anthropic tool result", is_error: false });
        return new Response(JSON.stringify({ content: [{ type: "text", text: "searched with Anthropic tools" }] }), { status: 200 });
      },
    });
    expect(result.text).toBe("searched with Anthropic tools");
    expect(result.toolResults?.[0]?.name).toBe("web_search");
    expect(statuses.some((status) => status.includes("Running web_search"))).toBe(true);
  });
  it("parses Anthropic-compatible responses", async () => {
    const result = await completeChat({ settings: { ...base, provider: "anthropic-compatible", route: "anthropic-v1" }, messages }, {
      fetch: async () => new Response(JSON.stringify({ content: [{ type: "text", text: "connected" }] }), { status: 200 }),
    });
    expect(result.text).toBe("connected");
  });

  it("propagates caller cancellation to the provider request", async () => {
    const controller = new AbortController();
    const pending = completeChat({ settings: base, messages, signal: controller.signal }, {
      fetch: async (_url, init) => new Promise((_resolve, reject) => {
        init?.signal?.addEventListener("abort", () => reject(Object.assign(new Error("aborted"), { name: "AbortError" })), { once: true });
      }),
    });
    controller.abort();
    await expect(pending).rejects.toThrow(/timed out|aborted/i);
  });

  it("passes caller cancellation into streaming provider requests", async () => {
    const controller = new AbortController();
    let receivedSignal: AbortSignal | null | undefined;
    const pending = completeChat({ settings: base, messages, signal: controller.signal, stream: true }, {
      fetch: async (_url, init) => {
        receivedSignal = init?.signal;
        return new Promise((_resolve, reject) => init?.signal?.addEventListener("abort", () => reject(Object.assign(new Error("aborted"), { name: "AbortError" })), { once: true }));
      },
    });
    controller.abort();
    await expect(pending).rejects.toThrow(/aborted|cancelled/i);
    expect(receivedSignal).not.toBeUndefined();
  });
});



