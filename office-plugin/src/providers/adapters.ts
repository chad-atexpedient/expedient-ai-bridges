import { summarizeAttachment } from "../lib/artifacts";
import { capabilityPrompt, isToolAvailable, toolDefinitionsForCapabilities } from "../lib/capabilities";
import { normalizeToolCall } from "../lib/tools";
import { runtimeInstructionBundle } from "../skills/registry";
import type { AttachmentContext, ChatMessage, CompletionRequest, CompletionResult, DocumentContext, ProviderSettings, ToolCallResult } from "../lib/types";

interface HttpClient {
  fetch(input: string, init?: RequestInit): Promise<Response>;
}

const REQUEST_TIMEOUT_MS = 120000;
const MAX_TOOL_ROUNDS = 12;
const defaultHttpClient: HttpClient = { fetch: (...args) => fetch(...args) };

async function fetchWithTimeout(client: HttpClient, input: string, init?: RequestInit, timeoutMs = REQUEST_TIMEOUT_MS) {
  const controller = new AbortController();
  let timedOut = false;
  const abort = () => controller.abort();
  init?.signal?.addEventListener("abort", abort, { once: true });
  const timer = globalThis.setTimeout(() => { timedOut = true; controller.abort(); }, timeoutMs);
  try {
    return await client.fetch(input, { ...init, signal: controller.signal });
  } catch (error: any) {
    if (error?.name === "AbortError" && timedOut && !init?.signal?.aborted) throw new Error("The provider request timed out while the add-in was still waiting for a response.");
    throw error;
  } finally {
    globalThis.clearTimeout(timer);
    init?.signal?.removeEventListener("abort", abort);
  }
}

function cleanBaseUrl(baseUrl: string) {
  return baseUrl.replace(/\/+$/, "");
}

function providerPath(settings: ProviderSettings, kind: "chat" | "messages") {
  if (settings.route === "custom") return "";
  if (kind === "messages") return "/messages";
  if (settings.route === "openwebui-api") return "/chat/completions";
  return "/chat/completions";
}

function providerUrl(settings: ProviderSettings, kind: "chat" | "messages") {
  const target = `${cleanBaseUrl(settings.baseUrl)}${providerPath(settings, kind)}`;
  if (!settings.useLocalProxy) return target;
  return `/api/proxy?target=${encodeURIComponent(target)}`;
}

function hostDisplayName(host?: DocumentContext["host"] | null) {
  if (host === "excel") return "Excel";
  if (host === "word") return "Word";
  if (host === "powerpoint") return "PowerPoint";
  return "Office preview/unknown host";
}

function contextPrompt(context?: DocumentContext | null, attachments: AttachmentContext[] = [], warning?: string, capabilities?: CompletionRequest["capabilities"]) {
  const parts: string[] = [];
  const runtimeBundle = runtimeInstructionBundle(context?.host ?? capabilities?.host ?? "unknown");
  if (runtimeBundle.prompt) parts.push("Agent and skill runtime instructions:\n" + runtimeBundle.prompt);
  if (context?.text) {
    parts.push([
      `Office host: ${context.host}`,
      `Selection: ${context.selectionLabel}`,
      `Document title: ${context.title || "Untitled"}`,
      "Relevant document context:",
      context.text,
    ].join("\n"));
  }
  if (warning) parts.push(`Office context warning: ${warning}`);
  const capabilityText = capabilityPrompt(capabilities);
  if (capabilityText) parts.push(capabilityText);
  if (attachments.length) parts.push(attachments.map((attachment) => summarizeAttachment(attachment.name, attachment.text)).join("\n\n"));
  return parts.length ? parts.join("\n\n") : null;
}


function runtimeUserPromptCard(context?: DocumentContext | null, warning?: string, capabilities?: CompletionRequest["capabilities"]) {
  const host = context?.host ?? capabilities?.host ?? "unknown";
  const availableTools = capabilities?.availableTools?.length ? capabilities.availableTools.slice(0, 36).join(", ") : "none reported";
  const notes = capabilities?.notes?.length ? capabilities.notes.slice(0, 4).join("; ") : "none";
  const runtimeBundle = runtimeInstructionBundle(host);
  const activeAgent = runtimeBundle.agent ? `${runtimeBundle.agent.name} (${runtimeBundle.agent.id})` : "none";
  const supportingAgents = runtimeBundle.supportingAgents.length ? runtimeBundle.supportingAgents.map((agent: { name: string; id: string }) => `${agent.name} (${agent.id})`).join(", ") : "none";
  const activeSkills = runtimeBundle.skills.length ? runtimeBundle.skills.map((skill: { name: string; id: string }) => `${skill.name} (${skill.id})`).join(", ") : "none";
  return [
    "Runtime context for this request:",
    `- Current Office surface: ${hostDisplayName(host)}. You are operating inside this host, not a generic browser chat.`,
    context?.title ? `- Current file/title: ${context.title}.` : "- Current file/title: unavailable or not yet read.",
    context?.selectionLabel ? `- Current selection/context label: ${context.selectionLabel}.` : "- Current selection/context label: unavailable.",
    warning ? `- Context warning: ${warning}` : "- Context warning: none.",
    `- Available native tools for this surface: ${availableTools}.`,
    `- Active agent: ${activeAgent}.`,
    `- Supporting agents: ${supportingAgents}.`,
    `- Active skills: ${activeSkills}.`,
    `- Capability notes: ${notes}.`,
    `- Required behavior: understand the user's request as a task in ${hostDisplayName(host)}; use native tools when available to make real changes; do not give manual instructions or placeholders when a tool or generated Office file can perform the task.`,
    "- Tool argument rules: use local A1 ranges like A1:C12 for Excel, slideNumber for human-facing PowerPoint slide numbers, standard chart/shape names, and positive dimensions.",
  ].join("\n");
}

function messagesWithRuntimePrompt(messages: ChatMessage[], context?: DocumentContext | null, warning?: string, capabilities?: CompletionRequest["capabilities"]) {
  const latestUserIndex = [...messages].map((message, index) => ({ message, index })).reverse().find((item) => item.message.role === "user")?.index;
  if (latestUserIndex === undefined) return messages;
  const card = runtimeUserPromptCard(context, warning, capabilities);
  return messages.map((message, index) => index === latestUserIndex
    ? { ...message, content: `${message.content}

---
${card}` }
    : message);
}
function hostOutputInstruction(host?: DocumentContext["host"]) {
  if (host === "excel") return "Ground yourself in the workbook before changing it: call excel_get_workbook_overview when the layout is unknown, excel_read_range to read actual cell values or formulas, and excel_search_workbook to locate data by text, value, or formula reference. Never guess a sheet name, address, or existing value. excel_write_range refuses to overwrite non-empty cells unless overwrite is true, so read the target range first and only pass overwrite=true when replacing that data is clearly what the user asked for. When the user asks to fill, update, calculate, format, chart, organize, clean data, split columns, deduplicate rows, append/combine sheets or tables, merge lookup columns by key, audit, review formulas, trace precedents/dependents, assess risk, or add images to a sheet, prefer using Excel tools directly, including excel_add_worksheet, excel_rename_worksheet, excel_delete_worksheet, excel_set_worksheet_visibility, excel_clear_range, excel_write_range, excel_import_context_table for converting Word/PowerPoint/M365/upload/web/chat context into a structured worksheet table, excel_set_formula, excel_clean_transform_range for Power Query-style cleanup into a new structured table, excel_combine_ranges for append/merge combine workflows, excel_audit_formulas for Formula Auditing-style risk review sheets, excel_map_formula_dependencies for precedent/dependent lineage maps, excel_create_chart, excel_create_pivottable for native PivotTables where available, excel_create_summary_table for pivot-style fallback summaries, excel_create_pivot_chart_report for grouped summary tables plus charts/PivotChart-style reports, excel_create_table, excel_format_range, excel_sort_range, excel_filter_range, excel_freeze_panes, excel_add_data_validation, excel_apply_conditional_format, excel_insert_image for uploaded or web images, excel_add_comment for review notes, excel_set_named_range for semantic workbook references, excel_protect_sheet for locking worksheets, and excel_set_page_layout for print/page setup, print areas, repeating headers, scaling, gridlines/headings, and centering. Use excel_generate_workbook_file when a standalone XLSX artifact with sheets, tables, validations, conditional formats, formulas, chart objects, or embedded images is better than live mutation. For cross-surface context imports into Excel, use excel_import_context_table after reading the relevant Word, PowerPoint, M365, uploaded, web, or chat context. For cleanup/import requests against an existing worksheet range, prefer excel_clean_transform_range before manual Text to Columns, Power Query, trim, blank-row, or duplicate-removal instructions; for combine requests, prefer excel_combine_ranges before manual append/merge/Paste/Power Query instructions. For analyst pivot or visual summary requests, prefer excel_create_pivot_chart_report when the user wants a charted grouped summary, prefer excel_create_pivottable if they specifically need a native PivotTable and it is available, and otherwise use excel_create_summary_table or excel_generate_workbook_file rather than manual PivotTable/PivotChart instructions. For workbook audit requests, use excel_audit_formulas and excel_map_formula_dependencies before giving manual formula-auditing instructions. Do not give manual chart/image/table/audit/cleanup/combine instructions when an Excel tool or generated workbook can do it. If you answer with data instead, return CSV, a Markdown table, or a JSON array of objects so it can be inserted as a real cell range.";
  if (host === "powerpoint") return "When the user asks for slide content, prefer using PowerPoint tools directly for simple live edits. Use powerpoint_audit_deck for deck QA, executive/design review, accessibility/design-readiness checks, slide-story review, and professional polish audits; use powerpoint_create_slides for slide text, powerpoint_add_textbox for positioned copy, powerpoint_add_table for slide tables, powerpoint_add_shape for simple visual structure, powerpoint_create_diagram for live editable process flows, cycles, roadmaps, timelines, workflows, and simple diagram/framework requests, powerpoint_arrange_shapes for alignment/distribution/z-order polish, powerpoint_group_shapes for grouping or ungrouping related objects when supported, powerpoint_set_slide_background for slide design, powerpoint_delete_slide and powerpoint_clear_slide for cleanup, powerpoint_insert_image for uploaded or public web images, powerpoint_duplicate_slide for template-like slide reuse, powerpoint_add_hyperlink_textbox for linked call-to-action text when supported, and powerpoint_generate_deck_file when live PowerPoint APIs cannot reliably create the requested deck artifact. For demos, prepared showcase decks, speaker notes, embedded images, backgrounds, designed shapes, table objects, clickable hyperlinks, chart objects, footer/date/slide-number chrome, or preservation of an attached/source PPTX template, prefer powerpoint_generate_deck_file only when the user asked for a downloadable artifact or live APIs cannot make the requested object. Do not call powerpoint_add_speaker_notes for generated/demo decks. If the user attached a PPTX and asks to use it as a template, match branding, preserve theme/layouts, or create a deck from that source, pass templateAssetName or templateAssetId into powerpoint_generate_deck_file; choose layoutName/layoutType per slide from the attached deck context when available; provide slide title, subtitle when useful, body, images, tables, notes, and editable charts. For generated charts, choose professional chartType, axis titles, valueFormat, dataLabels, colors, legendPosition, barDirection, and gridline visibility when useful; for comparisons across years, regions, products, or scenarios, use categories plus series instead of flattening everything into one single-series values list. Omit object coordinates when the template should control placement through image/table/chart placeholders. Never create picture/table/shape/chart placeholders if a real image asset, public imageUrl, generated deck, or web-search/fetch path can provide the object. Do not give manual deck-review/design-check instructions when powerpoint_audit_deck, powerpoint_create_diagram, or powerpoint_generate_deck_file can create a concrete review/artifact. If you answer with slide text instead, separate slide-ready sections with a line containing only --- slide.";
  if (host === "word") return "When the user asks for document content, prefer using Word tools directly. Use word_audit_document for document QA, accessibility review, structure/readability checks, and professional polish audits; use word_create_context_brief after reading Excel, PowerPoint, Microsoft 365, uploaded, web, or chat context when the user asks to turn that context into a memo, brief, pre-read, summary, or narrative Word artifact; use word_insert_heading for document structure, word_insert_table for tabular content, word_apply_style for styling, word_insert_page_break and word_insert_section_break for pagination/layout, word_set_header_footer for report headers/footers, word_find_replace for edits, word_insert_comment for review notes, word_insert_image for uploaded or public web images, word_insert_hyperlink for real links, word_format_selection for common formatting, word_insert_content_control for template/form fields and controlled placeholders, and word_generate_document_file when a standalone DOCX artifact with sections, real bulleted/numbered lists, tables, headers, footers, comments, tracked-change revision markup/redlines, footnotes, endnotes, clickable hyperlinks, table of contents, page layout, or embedded images is better than live mutation. If the user attaches a DOCX and asks to preserve its branding or template, pass templateAssetId or templateAssetName to word_generate_document_file. Do not give manual layout/image/TOC/footnote/endnote/redline/list/accessibility-check instructions when a Word tool or generated document can do it. If you answer with text instead, return clean prose with headings and lists.";
  return "When producing data for Office, use a structured format that can be inserted cleanly into the target document.";
}

function systemInstruction(context?: DocumentContext | null) {
  return [
    "You are an Office productivity assistant. Use the provided live document context and attached file context when relevant. Do not invent document contents.",
    "When tools are available, use native tool/function calls to search the web for current facts and to directly update Excel, Word, or PowerPoint when the user asks you to make changes. For image requests, first use uploaded image assets when provided; otherwise use web_image_search to find real imageUrl candidates, use web_image_import to validate/cache the chosen URL when practical, then use the assetId/assetName or imageUrl through native image insertion or generated Office file tools. Do not substitute placeholder text for a real image unless no image source is available and you clearly say what is missing.",
    "The app renders simple demo/showcase/tour requests locally in the chat pane. Call ctrl_create_demo_showcase only when the user explicitly asks to create a live Office sample in the file, downloadable demo files, or a full Excel/Word/PowerPoint demo suite. Use surface=current and mode=live for an explicit live file demo; use surface=all or mode=artifact for downloadable demo files.",
    "When the user asks for context from other Microsoft 365 files, first use m365_try_office_sso to reuse the profile already signed into Excel, Word, or PowerPoint, then use m365_auth_status, m365_search_files, and m365_read_file. If Office SSO is unavailable and a Microsoft device-login code is returned, show the verification URL and code clearly, then wait for the user to complete sign-in before continuing.",
    "Plan briefly, use the smallest useful number of tool calls, and stop calling tools once the requested Office change has been completed. For Office write tools, use simple normalized arguments: local A1 ranges like A1:C12 when sheetName is supplied, slideNumber for human-facing PowerPoint slide numbers, standard chart/shape names, and positive dimensions. If a live Office write tool returns an invalid-argument, missing-property, context.sync/load, unsupported, or capability error, retry at most once with simpler arguments; if it still fails, switch to the matching generated Office file tool (excel_generate_workbook_file, word_generate_document_file, or powerpoint_generate_deck_file) instead of placeholders or manual instructions. Do not keep calling the same failing live tool.",
    "After using Office write tools, summarize what you changed instead of repeating the full inserted content unless the user asks for it. If a generated Office file tool returns markdownLink, show that exact markdownLink verbatim; if it only returns downloadUrl, show that URL clearly as the created artifact.",
    hostOutputInstruction(context?.host),
  ].join(" ");
}

function toOpenAIMessages(messages: ChatMessage[], context?: DocumentContext | null, attachments: AttachmentContext[] = [], warning?: string, capabilities?: CompletionRequest["capabilities"]) {
  const prompt = contextPrompt(context, attachments, warning, capabilities);
  const apiMessages = messagesWithRuntimePrompt(messages, context, warning, capabilities).map((message) => ({ role: message.role, content: message.content }));
  if (prompt) {
    return [
      { role: "system" as const, content: systemInstruction(context) },
      { role: "system" as const, content: prompt },
      ...apiMessages,
    ];
  }
  return [{ role: "system" as const, content: systemInstruction(context) }, ...apiMessages];
}

function toAnthropicMessages(messages: ChatMessage[], context?: DocumentContext | null, attachments: AttachmentContext[] = [], warning?: string, capabilities?: CompletionRequest["capabilities"]) {
  const prompt = contextPrompt(context, attachments, warning, capabilities);
  const system = prompt
    ? `${systemInstruction(context)}\n\n${prompt}`
    : systemInstruction(context);
  return {
    system,
    messages: messagesWithRuntimePrompt(messages, context, warning, capabilities)
      .filter((message) => message.role !== "system")
      .map((message) => ({ role: message.role === "assistant" ? "assistant" : "user", content: message.content })),
  };
}

export function selectModel(settings: ProviderSettings, messages: ChatMessage[], context?: DocumentContext | null, attachments: AttachmentContext[] = []) {
  if (settings.modelMode === "manual") return settings.model;

  const [lightModel, standardModel, deepModel] = settings.autoModels;
  const latestUser = latestUserText(messages);
  const attachmentText = attachments.map((attachment) => attachment.text).join("\n");
  const combined = `${latestUser}\n${context?.text ?? ""}\n${attachmentText}`.toLowerCase();
  const contextLength = (context?.text?.length ?? 0) + attachmentText.length;

  const deepSignals = [
    "build", "create", "generate", "audit", "analyze", "strategy", "forecast", "model", "macro", "pivot", "deck", "presentation",
    "compare", "complex", "rewrite", "refactor", "formula", "financial", "multi-step", "workflow", "implementation",
  ];
  const standardSignals = ["summarize", "explain", "draft", "review", "clean", "format", "classify", "extract", "table"];

  if (contextLength > 5000 || deepSignals.some((signal) => combined.includes(signal))) return deepModel;
  if (contextLength > 1200 || standardSignals.some((signal) => combined.includes(signal))) return standardModel;
  return lightModel;
}

function withOptionalOpenAISettings(body: Record<string, unknown>, settings: ProviderSettings) {
  const next = { ...body };
  if (settings.temperature !== null) next.temperature = settings.temperature;
  if (settings.maxTokens !== null) next.max_tokens = settings.maxTokens;
  return next;
}

function withOptionalAnthropicSettings(body: Record<string, unknown>, settings: ProviderSettings) {
  const next = { ...body };
  if (settings.temperature !== null) next.temperature = settings.temperature;
  if (settings.maxTokens !== null) next.max_tokens = settings.maxTokens;
  return next;
}


function anthropicToolsForCapabilities(capabilities?: CompletionRequest["capabilities"]) {
  return toolDefinitionsForCapabilities(capabilities).map((tool: any) => ({
    name: tool.function.name,
    description: tool.function.description,
    input_schema: tool.function.parameters,
  }));
}

function anthropicText(content: any) {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  return content.map((part) => part?.type === "text" ? part.text ?? "" : "").join("").trim();
}

function anthropicToolUses(content: any) {
  if (!Array.isArray(content)) return [];
  return content.filter((part) => part?.type === "tool_use" && part?.id && part?.name);
}

function normalizeAnthropicToolCall(part: any, index: number) {
  return normalizeToolCall({ id: part.id || `anthropic-tool-${Date.now()}-${index}`, name: part.name, arguments: part.input ?? {} }, index);
}

function forceFinalAnthropicMessages(transcript: any[]) {
  return [
    ...transcript,
    {
      role: "user" as const,
      content: "Stop calling tools now. Produce the final user-facing answer. Summarize completed actions, mention any remaining uncertainty, and include the key result. Do not request more tools.",
    },
  ];
}

function latestUserText(messages: ChatMessage[]) {
  return [...messages].reverse().find((message) => message.role === "user")?.content ?? "";
}

function nativeToolNeeded(request: CompletionRequest) {
  const text = latestUserText(request.messages);
  if (/\b(image|images|photo|photos|picture|pictures|visual|visuals|graphic|graphics|logo|logos|background|thumbnail|thumbnails|dolphin|calf|calves|pod|pods|ocean|scene|scenes)\b/i.test(text)) return "web_image_search";
  if (/\b(search|web|look up|lookup|latest|current|recent|today|this month|this year|last\s+\d+\s+(?:days|weeks|months|years)|over the last|source|sources)\b/i.test(text)) return "web_search";
  return null;
}

function forceFinalMessages(transcript: any[]) {
  return [
    ...transcript,
    {
      role: "system" as const,
      content: "Stop calling tools now. Produce the final user-facing answer. Summarize completed actions, mention any remaining uncertainty, and include the key result. Do not request more tools.",
    },
  ];
}

function forceToolMessages(transcript: any[], toolName: string) {
  return [
    ...transcript,
    {
      role: "system" as const,
      content: `You must call the ${toolName} tool now. Choose concise arguments from the user's latest request. Do not answer directly before calling the tool.`,
    },
  ];
}

export function buildOpenAIRequest(
  settings: ProviderSettings,
  messages: ChatMessage[],
  context?: DocumentContext | null,
  contextWarning?: string,
  attachments: AttachmentContext[] = [],
  toolsEnabled = false,
  toolMessages?: any[],
  stream = false,
  forcedToolName?: string | null,
  capabilities?: CompletionRequest["capabilities"],
) {
  return {
    url: providerUrl(settings, "chat"),
    init: {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-provider-authorization": `Bearer ${settings.apiKey}`,
      },
      body: JSON.stringify(withOptionalOpenAISettings({
        model: selectModel(settings, messages, context, attachments),
        messages: toolMessages ?? toOpenAIMessages(messages, context, attachments, contextWarning, capabilities),
        ...(toolsEnabled ? { tools: toolDefinitionsForCapabilities(capabilities), tool_choice: forcedToolName ? { type: "function", function: { name: forcedToolName } } : "auto" } : {}),
        ...(stream ? { stream: true } : {}),
      }, settings)),
    } satisfies RequestInit,
  };
}

export function buildAnthropicRequest(
  settings: ProviderSettings,
  messages: ChatMessage[],
  context?: DocumentContext | null,
  contextWarning?: string,
  attachments: AttachmentContext[] = [],
  capabilities?: CompletionRequest["capabilities"],
  toolsEnabled = false,
  toolMessages?: any[],
) {
  const body = toAnthropicMessages(messages, context, attachments, contextWarning, capabilities);
  return {
    url: providerUrl(settings, "messages"),
    init: {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-provider-api-key": settings.apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(withOptionalAnthropicSettings({
        model: selectModel(settings, messages, context, attachments),
        system: body.system,
        messages: toolMessages ?? body.messages,
        ...(toolsEnabled ? { tools: anthropicToolsForCapabilities(capabilities) } : {}),
      }, settings)),
    } satisfies RequestInit,
  };
}

async function parseJsonResponse(response: Response) {
  const text = await response.text();
  let json: any = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }

  if (!response.ok) {
    const message = json?.error?.message || json?.message || json?.detail || text || `${response.status} ${response.statusText}`;
    throw new Error(typeof message === "string" ? message : JSON.stringify(message));
  }
  return json;
}

async function postBuiltRequest(url: string, init: RequestInit, client: HttpClient, signal?: AbortSignal) {
  return parseJsonResponse(await fetchWithTimeout(client, url, { ...init, signal: signal ?? init.signal }));
}

export async function completeChat(request: CompletionRequest, client: HttpClient = defaultHttpClient): Promise<CompletionResult> {
  if (!request.settings.apiKey.trim()) throw new Error("Add your API key in Settings first.");

  const built = request.settings.provider === "openai-compatible"
    ? buildOpenAIRequest(
        request.settings,
        request.messages,
        request.context,
        request.contextWarning,
        request.attachments,
        request.toolsEnabled,
        undefined,
        Boolean(request.stream && !request.toolsEnabled),
        null,
        request.capabilities,
      )
    : buildAnthropicRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, request.capabilities, request.toolsEnabled);

  if (request.stream && request.settings.provider === "openai-compatible" && !request.toolsEnabled) {
    return completeOpenAIStream(built.url, built.init, request, client);
  }

  request.onStatus?.(request.toolsEnabled ? "Contacting model with native tools enabled..." : "Contacting model...");
  const json = await postBuiltRequest(built.url, built.init, client, request.signal);

  if (request.settings.provider === "openai-compatible") {
    if (request.toolsEnabled && request.toolExecutor && Array.isArray(json?.choices?.[0]?.message?.tool_calls)) {
      return completeOpenAIToolLoop(request, json, built.url, client);
    }
    const requiredTool = request.toolsEnabled && request.toolExecutor ? nativeToolNeeded(request) : null;
    if (requiredTool && isToolAvailable(request.capabilities, requiredTool)) {
      const nudged = await requestNativeToolCall(request, built.url, client, requiredTool);
      if (nudged) return nudged;
    }
    const output = json?.choices?.[0]?.message?.content;
    if (!output) throw new Error("The OpenAI-compatible endpoint returned no assistant content.");
    return { text: output, raw: json };
  }

  if (request.toolsEnabled && request.toolExecutor && anthropicToolUses(json?.content).length) {
    return completeAnthropicToolLoop(request, json, built.url, client);
  }

  const content = anthropicText(json?.content);
  if (!content) throw new Error("The Anthropic-compatible endpoint returned no assistant content.");
  return { text: content, raw: json };
}


async function requestNativeToolCall(request: CompletionRequest, url: string, client: HttpClient, toolName: string): Promise<CompletionResult | null> {
  request.onStatus?.(`Requesting native ${toolName} tool call...`);
  const transcript = forceToolMessages(toOpenAIMessages(request.messages, request.context, request.attachments, request.contextWarning, request.capabilities), toolName);
  const built = buildOpenAIRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, true, transcript, false, toolName, request.capabilities);
  const json = await postBuiltRequest(url, built.init, client, request.signal);
  if (!Array.isArray(json?.choices?.[0]?.message?.tool_calls)) return null;
  return completeOpenAIToolLoop(request, json, url, client, transcript);
}

async function completeOpenAIToolLoop(request: CompletionRequest, firstJson: any, url: string, client: HttpClient, existingTranscript?: any[]): Promise<CompletionResult> {
  const transcript: any[] = existingTranscript ?? toOpenAIMessages(request.messages, request.context, request.attachments, request.contextWarning, request.capabilities);
  let currentJson = firstJson;
  const toolResults: ToolCallResult[] = [];

  for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
    if (request.signal?.aborted) throw new DOMException("The request was cancelled.", "AbortError");
    request.onStatus?.(`Thinking through tool round ${round + 1}...`);
    const assistantMessage = currentJson?.choices?.[0]?.message;
    const toolCalls = Array.isArray(assistantMessage?.tool_calls) ? assistantMessage.tool_calls : [];
    if (!toolCalls.length) {
      const output = assistantMessage?.content;
      if (!output) {
        const fallbackText = generatedArtifactSummary(toolResults);
        if (fallbackText) return { text: fallbackText, raw: currentJson, toolResults };
        throw new Error("The OpenAI-compatible endpoint returned no assistant content after tool calls.");
      }
      request.onToken?.(output);
      return { text: output, raw: currentJson, toolResults };
    }

    transcript.push(assistantMessage);
    for (let index = 0; index < toolCalls.length; index += 1) {
      const call = normalizeToolCall(toolCalls[index], index);
      request.onStatus?.(`Running ${call.name}...`);
      request.onToolStart?.(call);
      if (request.signal?.aborted) throw new DOMException("The request was cancelled.", "AbortError");
      const result = await request.toolExecutor!(call);
      toolResults.push(result);
      request.onToolResult?.(result);
      transcript.push({ role: "tool", tool_call_id: call.id, content: result.content });
    }

    const nextBuilt = buildOpenAIRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, true, transcript, false, null, request.capabilities);
    request.onStatus?.("Sending tool results back to the model...");
    currentJson = await postBuiltRequest(url, nextBuilt.init, client, request.signal);
  }

  request.onStatus?.("Asking the model for a final summary...");
  const finalBuilt = buildOpenAIRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, false, forceFinalMessages(transcript), false, null, request.capabilities);
  const finalJson = await postBuiltRequest(url, finalBuilt.init, client, request.signal);
  const output = finalJson?.choices?.[0]?.message?.content;
  if (output) request.onToken?.(output);
  const fallbackText = generatedArtifactSummary(toolResults);
  return {
    text: output || fallbackText || `I ran ${toolResults.length} tool call${toolResults.length === 1 ? "" : "s"}. The model did not provide a final summary, but the task list shows what completed.`,
    raw: finalJson,
    toolResults,
  };
}

function generatedArtifactSummary(toolResults: ToolCallResult[]) {
  const links: string[] = [];
  for (const result of toolResults) {
    if (!result.ok) continue;
    const parsed = safeJson(result.content);
    const candidates = Array.isArray(parsed?.generated) ? parsed.generated.map((item: any) => item?.result) : [parsed];
    for (const candidate of candidates) {
      const markdownLink = typeof candidate?.markdownLink === "string" ? candidate.markdownLink.trim() : "";
      if (markdownLink) links.push(markdownLink);
    }
  }
  if (!links.length) return "";
  return `Created the requested Office artifact${links.length === 1 ? "" : "s"}.\n\n${links.join("\n")}`;
}

function safeJson(text: string) {
  try { return JSON.parse(text); } catch { return null; }
}

async function completeAnthropicToolLoop(request: CompletionRequest, firstJson: any, url: string, client: HttpClient): Promise<CompletionResult> {
  const base = toAnthropicMessages(request.messages, request.context, request.attachments, request.contextWarning, request.capabilities);
  const transcript: any[] = [...base.messages];
  let currentJson = firstJson;
  const toolResults: ToolCallResult[] = [];

  for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
    if (request.signal?.aborted) throw new DOMException("The request was cancelled.", "AbortError");
    request.onStatus?.(`Thinking through Anthropic tool round ${round + 1}...`);
    const content = Array.isArray(currentJson?.content) ? currentJson.content : [];
    const toolUses = anthropicToolUses(content);
    if (!toolUses.length) {
      const output = anthropicText(content);
      if (!output) {
        const fallbackText = generatedArtifactSummary(toolResults);
        if (fallbackText) return { text: fallbackText, raw: currentJson, toolResults };
        throw new Error("The Anthropic-compatible endpoint returned no assistant content after tool calls.");
      }
      request.onToken?.(output);
      return { text: output, raw: currentJson, toolResults };
    }

    transcript.push({ role: "assistant", content });
    const resultContent: any[] = [];
    for (let index = 0; index < toolUses.length; index += 1) {
      const call = normalizeAnthropicToolCall(toolUses[index], index);
      request.onStatus?.(`Running ${call.name}...`);
      request.onToolStart?.(call);
      if (request.signal?.aborted) throw new DOMException("The request was cancelled.", "AbortError");
      const result = await request.toolExecutor!(call);
      toolResults.push(result);
      request.onToolResult?.(result);
      resultContent.push({ type: "tool_result", tool_use_id: toolUses[index].id, content: result.content, is_error: !result.ok });
    }
    transcript.push({ role: "user", content: resultContent });

    const nextBuilt = buildAnthropicRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, request.capabilities, true, transcript);
    request.onStatus?.("Sending tool results back to the model...");
    currentJson = await postBuiltRequest(url, nextBuilt.init, client, request.signal);
  }

  request.onStatus?.("Asking the model for a final summary...");
  const finalBuilt = buildAnthropicRequest(request.settings, request.messages, request.context, request.contextWarning, request.attachments, request.capabilities, false, forceFinalAnthropicMessages(transcript));
  const finalJson = await postBuiltRequest(url, finalBuilt.init, client, request.signal);
  const output = anthropicText(finalJson?.content);
  if (output) request.onToken?.(output);
  const fallbackText = generatedArtifactSummary(toolResults);
  return {
    text: output || fallbackText || `I ran ${toolResults.length} tool call${toolResults.length === 1 ? "" : "s"}. The model did not provide a final summary, but the task list shows what completed.`,
    raw: finalJson,
    toolResults,
  };
}

async function completeOpenAIStream(url: string, init: RequestInit, request: CompletionRequest, client: HttpClient): Promise<CompletionResult> {
  request.onStatus?.("Opening streaming response...");
  const response = await fetchWithTimeout(client, url, { ...init, signal: request.signal ?? init.signal }, REQUEST_TIMEOUT_MS);
  if (!response.ok) await parseJsonResponse(response);

  if (!response.body?.getReader) {
    const json = await parseJsonResponse(response);
    const output = json?.choices?.[0]?.message?.content || "";
    if (output) request.onToken?.(output);
    return { text: output, raw: json };
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let output = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const events = buffer.split("\n\n");
    buffer = events.pop() || "";

    for (const event of events) {
      for (const line of event.split("\n")) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("data:")) continue;
        const data = trimmed.slice(5).trim();
        if (!data || data === "[DONE]") continue;
        try {
          const json = JSON.parse(data);
          const token = json?.choices?.[0]?.delta?.content || "";
          if (token) {
            output += token;
            request.onToken?.(token);
          }
        } catch {
          // Ignore malformed streaming chunks from permissive gateways.
        }
      }
    }
  }

  if (!output) throw new Error("The OpenAI-compatible endpoint streamed no assistant content.");
  return { text: output };
}

export async function testProvider(settings: ProviderSettings, client?: HttpClient) {
  return completeChat({
    settings: { ...settings, maxTokens: null, temperature: null },
    messages: [{ id: "test", role: "user", content: "Reply with exactly: connected", timestamp: new Date().toISOString() }],
  }, client);
}




