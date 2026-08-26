# Office API capability hardening

This add-in treats Office APIs as runtime capabilities, not assumptions baked into the prompt.

## Current contract

On load, and again before each chat request, the task pane builds an Office capability snapshot from the active host and available Office.js APIs. That snapshot records:

- the current host: Excel, Word, PowerPoint, or preview;
- the Office platform reported by Office.js;
- detected requirement set support;
- model tools that are available right now;
- model tools that are unavailable, with a user-readable reason.

The model receives only the tools in the available list. If a stale or manually crafted tool call still asks for an unavailable action, the add-in blocks it before execution and sends the model a clear tool result explaining why.

Because BYOK gateways and model adapters do not all preserve or weight system messages the same way, every chat request also appends a compact runtime card to the latest user message. That card repeats the active Office surface, current file/title, selection/context label, available native tools, capability notes, and the no-placeholder/tool-first operating rules. This is intentional redundancy: Anthropic-compatible and OpenAI-compatible routes should both see that they are operating inside Excel, Word, or PowerPoint even if a gateway shim weakens hidden system context.

## Why this matters long term

Microsoft can add Office.js APIs, change rollout timing by channel, or make features available in one host/runtime before another. The add-in should adopt newer APIs by adding them to capability detection first, then exposing them as model tools only when the runtime proves support.

This keeps three things true:

1. New Office functionality can be enabled quickly without weakening older Office clients.
2. The model does not hallucinate Excel tools inside Word or PowerPoint tools inside Excel.
3. Production support gets actionable failure messages instead of generic stalled chats.

## Adding a new native Office tool

Use this path for future full-ability coverage:

1. Add the tool schema in `src/lib/tools.ts`.
2. Add runtime availability checks in `src/lib/capabilities.ts`.
3. Implement the Office.js action in `src/office/host.ts`.
4. Add provider/capability tests so the tool is only exposed in supported hosts.
5. Prefer direct Office mutation tools over manual user instructions whenever Office.js supports the action.

Excel chart note: `excel_create_chart` should place charts with cell anchors such as `startCell: "E2"` and `endCell: "M20"`. Office.js `chart.setPosition()` expects range/cell anchors, not raw point coordinates. Numeric `top`, `left`, `width`, and `height` can be applied as optional chart properties after anchoring.

Excel PivotTable note: `excel_create_pivottable` is exposed only when capability detection sees PivotTable-related Excel APIs. It accepts source range, destination sheet/cell, row/column/filter fields, value fields, and aggregation hints. If the installed Office runtime does not expose native PivotTable creation, the model should use `excel_create_summary_table` for a live pivot-style table or `excel_generate_workbook_file` for a generated workbook artifact.

Excel cross-surface import note: `excel_import_context_table` converts readable context from Word, PowerPoint, Microsoft 365 files, uploaded files, web results, or chat into a live Excel worksheet table. It detects markdown/CSV/TSV/JSON-like tables, key-value notes, bullets, and plain lines, adds a source column by default, and treats cells beginning with `=`, `+`, `-`, or `@` as text unless the caller explicitly enables formula preservation. Use it after `office_read_context`, `m365_read_file`, uploaded file context, or web fetch/search when the user asks to put external context into Excel.


## Current professional authoring baseline

The native tool baseline now covers more than simple insertion:

- Excel: worksheets, worksheet rename/delete/visibility, range clearing/writes, structured context imports from Word/PowerPoint/M365/uploads/web/chat, formulas, charts, native PivotTables when the runtime exposes PivotTable APIs, pivot-style summary-table fallbacks, structured tables, range formatting, sorting, filters, freeze panes, dropdown validation, conditional formatting, comments/notes where supported, named ranges, worksheet protection, page layout/print setup including print area, repeating titles, scaling, gridlines/headings, centering, and image insertion.
- Word: text, headings, tables, context briefs from Excel/PowerPoint/M365/uploads/web/chat, style application, font/selection formatting, page breaks, section breaks, primary header/footer updates, find/replace, comments when supported, hyperlinks where supported, guarded content controls/template fields, and image insertion.
- PowerPoint: slide creation, slide duplication where supported, positioned text boxes, hyperlink text boxes where supported, table grids/native tables when exposed, shapes, alignment/distribution/z-order arrangement, guarded shape grouping/ungrouping when exposed by the PowerPoint runtime, slide background color/image fallback, slide deletion/clearing when supported, speaker notes when supported, image insertion, and generated deck file creation.

Word content-control note: `word_insert_content_control` only reports success for native content controls when the Word runtime exposes content-control insertion methods. If those APIs are unavailable, the tool returns a clear limitation and the model should use `word_generate_document_file` for template-style DOCX generation.

Word cross-surface brief note: `word_create_context_brief` converts readable context from Excel, PowerPoint, Microsoft 365 files, uploaded files, web results, or chat into a live Word memo/pre-read structure. It inserts a heading, source/audience note, executive summary, key-point table, and an extracted source table when table-like data is detected. Use it after `office_read_context`, `m365_read_file`, uploaded file context, or web fetch/search when the user asks to turn external context into a Word brief instead of raw pasted text.

This is still not full ribbon parity. The live Office.js runtime remains the gate for direct mutation. For capabilities Office.js does not expose reliably, the production path should use server-side Open XML generation/manipulation, Microsoft Graph file operations, or a clear fallback that preserves user intent without pretending an object was inserted.

See `office-capability-roadmap.md` for the coverage matrix and next implementation priorities.


## Generated Office file lane

The add-in now has a generated Office file lane for cases where live Office.js APIs are too limited or inconsistent.

- `powerpoint_generate_deck_file` creates a downloadable `.pptx` through `/api/generated/pptx`.
- `word_generate_document_file` creates a downloadable `.docx` through `/api/generated/docx`.
- `excel_generate_workbook_file` creates a downloadable `.xlsx` through `/api/generated/xlsx`.
- The add-in renders simple demo/showcase/tour requests directly in the chat pane as a lightweight feature card. `ctrl_create_demo_showcase` is reserved for explicit live Office samples or downloadable artifact demos.
- Generated decks support slide title/body/subtitle text, speaker notes, background colors, editable footer/date/classification/slide-number text boxes, real geometric shapes, real table graphic frames with professional styling controls, clickable hyperlink text boxes, editable chart parts, real embedded images from uploaded asset names/ids, public image URLs, or direct base64 data, image fit/fill/crop controls for brand-safe placement, and template-preserving generation from an uploaded/source `.pptx` asset.
- Generated documents support title, table-of-contents field, document-level and section-specific page layout/orientation/margins, real Word columns, sections, headings, paragraphs, real bulleted and numbered list parts (`word/numbering.xml`), tables, real Word caption fields/bookmarks (`SEQ`) and cross-reference fields (`REF`), headers, footers, hyperlink relationships, footnote and endnote parts, comment parts, tracked-change style revision markup (`w:ins`/`w:del` plus `w:trackRevisions` settings), and real embedded images. Word may need to refresh fields after opening to populate generated TOC page numbers and field results.
- Generated workbooks support worksheets, typed values, inline strings, formulas beginning with `=`, styles, merged cells, column widths/visibility, row heights/hidden rows, plain range AutoFilters, structured table parts, dropdown validations, conditional-format scaffolding, workbook/sheet-scoped named ranges, worksheet view/page-layout settings, legacy cell comments/notes for review handoff, clickable external hyperlinks, workbook document properties (`docProps/core.xml` and `docProps/app.xml`), chart parts, and real embedded worksheet images.
- The server stores generated files under the local add-in generated-artifacts directory and returns a `downloadUrl` such as `/api/generated/files/<id>.pptx`.
- Generated Office files can be inspected by the same rich Open XML context extractor, which proves text, notes where applicable, media parts, and image relationships exist in the package.

This lane is the foundation for capabilities that do not fit live Office.js well: guaranteed speaker notes, template-aware deck generation, richer chart/object packaging, Word report/template exports with headers/footers/comments, Excel workbook exports with tables/validations/conditional formats, and cross-host transform workflows.

Prepared demo note: vague demo/tour/showcase requests are handled locally by the chat UI so the user immediately sees a native add-in-style feature card. From that card, the user can choose an explicit live Office sample (`surface=current`, `mode=live`) or downloadable generated sample files (`surface=all`, `mode=artifact`). This keeps the default demo simple and non-destructive while preserving the richer live/file demo lanes when requested.

Generated artifact handoff note: generated-file tools and the prepared demo tool include a `markdownLink` field when the server returns a `downloadUrl`. The assistant should display that field exactly, with no space between `]` and `(`, so the app renders a clean clickable artifact link.

Generated PowerPoint charts currently support bar/column, horizontal bar, line, area, pie, doughnut, scatter, and combo chart parts; single-series and multi-series category data; numeric scatter X/Y point series; combo clustered-column plus line series in one editable chart part; title, axis titles, value number format, optional value labels, major gridline toggling, legend position, simple per-series color styling, and doughnut hole size. Each generated chart part includes a relationship to an embedded `.xlsx` chart-data workbook under `ppt/embeddings/`, and the chart series point at workbook ranges (`strRef`/`numRef`) while retaining cached values for immediate display. Category charts use a `Chart Data` sheet with category labels in column A and series values in adjacent columns; scatter charts use paired X/Y columns per series. These are still Open XML chart parts, not screenshots, so PowerPoint can treat them as editable chart objects.

Generated Excel chart note: `excel_generate_workbook_file` can package self-contained bar/column, horizontal bar, line, area, pie, doughnut, scatter, and combo chart parts anchored through worksheet drawings. These chart parts intentionally remain self-contained chart caches unless a future workbook-relationship lane is added for Excel chart data packages. Do not add PowerPoint-only `externalData` relationships to `xl/charts/chart*.xml`; Excel-generated charts should avoid dangling relationship markers that can trigger invalid-package behavior.

Generated Excel metadata note: `excel_generate_workbook_file` can set professional document properties such as title, subject, creator/author, keywords, description, category, company, and manager. This matters for handoff, Microsoft 365 search, records/governance workflows, and distinguishing generated deliverables from scratch/test workbooks.

Generated Excel page-layout note: generated worksheets can carry frozen panes, gridline/headings visibility, zoom scale, print options, margins, paper size/orientation, scaling/fit-to-page, black-and-white/draft print flags, print area, and repeating title rows/columns. Print area and print titles are stored as workbook-level defined names, matching how Excel records those Page Layout toolbar settings.

Generated Excel protection note: generated worksheets can include `sheetProtection` with optional permissions for formatting cells, sorting, and autofiltering, plus the legacy Excel sheet-protection password hash when a password is supplied. This is accidental-edit prevention and workbook handoff polish, not encryption or strong protection for sensitive data.

Generated Excel named-range note: `excel_generate_workbook_file` can create workbook-scoped and sheet-scoped defined names. Use these for professional semantic formulas, reusable model inputs, print/model sections, and downstream automation rather than hard-coding raw coordinates everywhere. Names are sanitized to valid Excel defined-name characters when generated by the model.

Generated Excel comments note: generated worksheets can include legacy cell comments/notes with authors and optional visibility. The package includes both `xl/comments/commentN.xml` and a VML drawing relationship because that is the broadly compatible Open XML form for Excel cell notes. Uploaded/generated workbook context extraction also summarizes Excel comment parts so review notes remain visible to the model later.

Generated Excel hyperlink note: generated worksheets can include real clickable external hyperlinks with display text and screen tips. Links are represented through worksheet `<hyperlinks>` entries plus `sheetN.xml.rels` hyperlink relationships with `TargetMode="External"`, and uploaded workbook context extraction summarizes those hyperlink targets for future grounding.

Generated Excel layout note: generated worksheets can include real merged cell ranges, explicit column widths/hidden states, and row heights/hidden rows. Use these for report title bands, executive-summary sections, readable source columns, grouped detail sections, and print-ready layouts instead of fake spacer text or empty columns.

Generated Excel filter note: generated worksheets can include a plain range AutoFilter such as `A1:D100` when the sheet itself needs filter dropdowns. Structured tables already emit their own table-scoped AutoFilter, so use the sheet-level `autoFilter` option for non-table report ranges and raw data extracts.

PowerPoint note: `powerpoint_group_shapes` is intentionally guarded at execution because PowerPoint grouping APIs are not consistently exposed to task-pane add-ins. It should be used immediately when the runtime exposes `group`, `groupShapes`, `addGroup`, `createGroup`, or shape ungroup methods, and otherwise should return a clear limitation.

PowerPoint note: `powerpoint_generate_deck_file` must remain available in PowerPoint capability snapshots. It is the no-placeholder escape hatch when live PowerPoint APIs cannot place a real picture, table, shape, chart, hyperlink, background, or notes object but the user still needs a usable `.pptx` with embedded media and real slide objects.

PowerPoint speaker-notes note: live task-pane notes APIs are not reliably exposed across PowerPoint runtimes, so `powerpoint_add_speaker_notes` is hidden from normal capability snapshots unless a future runtime-specific detector proves support. For demos, prepared showcase decks, or any deck where notes must be guaranteed, use `powerpoint_generate_deck_file`; it writes speaker notes into the generated `.pptx` package directly.

Live Office failure policy: if a native write tool returns an invalid argument, missing property/load/context.sync, unsupported API, or runtime capability error, the model should retry at most once with simpler normalized arguments. After that, it should switch to the matching generated Office artifact lane (`excel_generate_workbook_file`, `word_generate_document_file`, or `powerpoint_generate_deck_file`) rather than looping on the same failing call or giving manual toolbar instructions. Tool failure payloads now include this recovery guidance plus redacted arguments so the model can self-correct without exposing keys or large base64 values.

Template preservation: when the user attaches a `.pptx`, the task pane stores it as a temporary open-pane template asset. The model can call `powerpoint_generate_deck_file` with `templateAssetId`, `templateAssetName`, or `templateBase64`. The server clones the source package, preserves non-slide template parts such as themes, slide masters, layouts, fonts, relationships, and reusable media, then replaces/adds generated slide XML linked to a source layout. Each generated slide can request `layoutName`, `layoutType`, or `layoutIndex`, so a single generated deck can use title, section, content, chart, table, or picture layouts from the same template. If the selected layout exposes title, body/content, or subtitle placeholders, generated slides fill those placeholders while keeping the placeholder geometry and relationship to the source layout. If it exposes picture/image, chart, or table placeholders, generated images, charts, and tables without explicit coordinates use those placeholder boxes as their placement targets. Current limitation: layout matching is metadata/name/type based; it does not yet infer brand-specific layout intent from visual design alone.

## Production hardening path

The XML manifest is still the correct Office add-in contract. Production should not rely on local XML sideloading per user; it should use Microsoft 365 Centralized Deployment or another managed Microsoft-approved deployment channel.

Recommended enterprise path:

1. Host this app on a stable HTTPS domain with a production certificate.
2. Generate a production manifest that points to that domain.
3. Deploy the manifest through Microsoft 365 Admin Center.
4. Keep Excel, Word, and PowerPoint on one shared settings contract. The current `/api/settings` endpoint is the seam: local/dev can store a shared settings blob, while production should resolve it from authenticated user or tenant state.
5. Move from local browser-stored BYOK secrets to MSAL/Entra sign-in plus a backend token/key service.
6. Store provider credentials in a customer-controlled vault and enforce provider/image/web-fetch allowlists at the backend. Production API access is fail-closed until the organization-specific bearer-token verifier and tenant policy layer are configured.
7. Keep the runtime capability snapshot enabled so Office channel differences degrade cleanly.

## Shared branding/key settings

Excel, Word, and PowerPoint should not maintain separate branding and provider settings. The app now loads and saves settings through a same-origin shared settings API first, with local browser storage only as a fallback mirror.

For development, `/api/settings` stores one shared settings blob for the local add-in service so all three hosts converge on the same product name, colors, logo, provider route, model defaults, and BYOK key.

For production, keep the client contract but replace the storage implementation with:

- MSAL/Entra authentication;
- tenant/user settings lookup;
- secret material stored server-side in a vault;
- optional policy controls for approved provider endpoints and model lists.

The task pane should still receive the same effective settings shape so the Excel, Word, and PowerPoint UX stays unified.

## Microsoft 365 delegated context

The add-in can now expose Microsoft 365 context tools to the model:

- `m365_try_office_sso` tries to reuse the Microsoft 365 profile already signed into Excel, Word, or PowerPoint through Office SSO.
- `m365_auth_status` checks whether delegated Graph access is ready and returns device-login instructions when it is not.
- `m365_search_files` searches the signed-in user's OneDrive/SharePoint files through Microsoft Graph.
- `m365_read_file` reads text context from a selected Microsoft 365 file and supports best-effort extraction from `.docx`, `.pptx`, and `.xlsx` Open XML files.

Local/development flow:

1. Create or reuse an Entra app registration for the add-in.
2. Replace the placeholder `WebApplicationInfo` id/resource through the production packaging inputs; `tools/make-production-manifest.mjs` requires HTTPS plus real `MSAL_CLIENT_ID` and `OFFICE_SSO_RESOURCE` values.
3. Set `MSAL_CLIENT_ID` on the add-in server. Optionally set `MSAL_TENANT_ID`; otherwise it uses `common`. If no client id is configured, the add-in returns placeholder setup guidance instead of guessing an app id.
4. Restart the add-in server and reload the Office add-in.
5. Ask the add-in to use a Microsoft 365 file. It first tries the signed-in Office profile via Office SSO.
6. If Office SSO is unavailable or Graph rejects the token, configure backend On-Behalf-Of exchange with `MSAL_CLIENT_SECRET`, or use device login as the fallback.
7. If no delegated token is cached, the fallback tool returns `https://microsoft.com/devicelogin` and a user code. Complete device login in the browser, then ask the add-in to continue.

Office SSO notes:

- Office does not hand arbitrary Microsoft 365 tokens to task panes automatically. The supported path is `OfficeRuntime.auth.getAccessToken()`/`Office.auth.getAccessToken()` plus `WebApplicationInfo` in the manifest.
- The server accepts the Office SSO token at `/api/m365/sso`.
- If `MSAL_CLIENT_SECRET` is configured, the server attempts OAuth On-Behalf-Of exchange for Microsoft Graph scopes.
- If OBO is not configured, the server may use a direct Office SSO token only in controlled development flows when Graph accepts that token's audience/scopes; production requires the approved OBO/session design.

Development can also set `GRAPH_ACCESS_TOKEN` to an already delegated Graph access token. That is useful when a trusted local environment already has delegated Microsoft 365 access, but it should not become the production architecture.

If a trusted local M365 tool already maintains a compatible `m365_token.json` file, set `M365_COMPAT_TOKEN_PATH` to that file path. The add-in server will import it into its own cache when possible. This is intentionally opt-in; the add-in should not silently scrape MCP or another tool's credential store.

Context reading is intentionally broad. This applies to both Microsoft 365 Graph files and user-uploaded files in the task pane. When clean extraction is available, the reader returns text from Office Open XML files (`.docx`, `.pptx`, `.xlsx`) or normal text-like files. When clean extraction is not available, it still returns bounded raw context: UTF-8 text if the bytes look textual, otherwise a base64 sample plus file metadata. That lets the model identify/translate custom formats or explain the converter needed instead of treating the file as unreadable.

PowerPoint and Word upload note: users should be able to attach PowerPoint and Word files for review, template generation, summarization, or transformation. The upload reader routes non-text files through `/api/file-context`, which uses the same extraction/fallback posture as the M365 file reader.

## Image and visual asset insertion

The add-in now treats images as first-class task-pane assets instead of only text context.

Supported paths:

- User uploads an image file (`.png`, `.jpg`, `.jpeg`, `.gif`, `.webp`, `.bmp`). The task pane keeps it in an in-memory asset registry for the current open pane session.
- The model can insert that image by `assetId` or filename with `excel_insert_image`, `word_insert_image`, or `powerpoint_insert_image`.
- If the user asks for topical visuals but does not provide images, the model can call `web_image_search` to get public image candidates with `imageUrl`, source/page URL, dimensions, and attribution hints.
- The model can pass a public `imageUrl` directly, or call `web_image_import` to fetch, validate, and cache that URL as a reusable in-pane asset first. The same-origin server fetches the image through `/api/image-asset`, validates type/size, converts it to base64, and returns an Office-friendly payload. This avoids browser/Office task-pane CORS problems.
- Generated `.pptx`, `.docx`, and `.xlsx` artifacts also embed uploaded or web images as local Open XML media parts. This is intentional: linked web images are more likely to break under tenant/network policy or disappear later.

PowerPoint no-placeholder rule: for image-heavy decks, the preferred sequence is uploaded image asset lookup or `web_image_search` -> `web_image_import`, then `powerpoint_insert_image` if the live runtime supports it, otherwise `powerpoint_generate_deck_file` with `assetId`, `assetName`, or `imageUrl` values so the downloadable `.pptx` contains real embedded media. Placeholder text such as "add a dolphin photo here" is only acceptable when no image source is available or policy blocks image use.

Runtime caveat: Office image APIs vary by host, platform, and Office channel. The tool definitions are available by host, but execution still checks for the actual Office method (`worksheet.shapes.addImage`, `insertInlinePictureFromBase64`, or PowerPoint slide image APIs). If the installed Office runtime does not expose the method, the tool returns a clear blocked reason instead of inserting placeholders.

Generated PowerPoint image layout note: `powerpoint_generate_deck_file` supports `fit`/`sizing` values of `fit`/`contain` for aspect-preserving letterbox placement, `fill`/`cover` for aspect-preserving center-crop placement, `stretch` for exact-box placement, and explicit crop percentages via `crop` or `cropLeft`/`cropRight`/`cropTop`/`cropBottom`. Web-image imports preserve detected source pixel dimensions when possible so generated decks can compute those placements without model guesswork.

Generated PowerPoint table note: `powerpoint_generate_deck_file` creates real DrawingML table graphic frames and now supports header fill/text color, body and banded-row fills, borders, font sizes, bolding, horizontal/vertical alignment, row heights, column widths, and per-cell overrides for highlighted values. Use this for executive KPI tables, comparison matrices, risk registers, and appendix tables instead of screenshotting or drawing rectangles around text.

Generated PowerPoint deck-chrome note: `powerpoint_generate_deck_file` supports global and per-slide `footer`/`footerText`, `dateText`/`date`, `confidentialityLabel`, `showSlideNumber`/`showSlideNumbers`, `slideNumberFormat`, `footerColor`, and `footerFontSize`. These render as normal editable text boxes on generated slides, including template-preserving decks, so board packs and client decks can carry deliverable-grade footer/date/classification/numbering polish without manual PowerPoint steps.

Production hardening targets:

- Add image search/license policy and source allowlists before encouraging public web image use. `web_image_import` is the intended enforcement point for license metadata, tenant allowlists, source reputation, and malware scanning.
- Add MIME sniffing, malware scanning, and tenant size limits on `/api/image-asset`.
- Add optional image normalization/cropping for live Office insertion where host runtimes expose reliable APIs; generated PowerPoint decks already have package-level fit/fill/crop controls.
- Persist approved assets server-side only when the user or tenant policy explicitly allows it.


## Rich uploaded/M365 Office file context

Uploaded and Microsoft 365 Office files now provide richer Open XML context instead of only plain body text:

- PowerPoint `.pptx`: slide text, object counts, slide relationships for images/layouts/charts/hyperlinks, speaker-note text when present, theme/master/layout counts, layout names/types/placeholders, and embedded media names.
- Word `.docx`: document body, style ids, comments, headers, footers, and embedded media names.
- Excel `.xlsx`: workbook structure, worksheet row/cell/filter/merge summaries, table names/ranges, shared strings, comment text, hyperlink targets, chart part names, and embedded media names.

This improves template/review workflows: the model can see whether an uploaded deck has themes, layouts, embedded images, notes, tables/charts, or brand assets before deciding how to edit or recreate content.

This is still context extraction, not guaranteed live mutation of every Open XML feature. The production generation lane should add server-side `.pptx`, `.docx`, and `.xlsx` creation/editing endpoints for capabilities that Office.js does not expose consistently.

## Session chat memory

The task pane supports natural new chats without losing grounding. When the user starts a new chat, the add-in compacts the prior chat into a temporary session-memory record containing:

- a short title;
- a compact transcript summary;
- source host/document hints;
- extracted keywords for retrieval.

While the file/task pane remains open, future prompts rank these memory records against the current user request and Office context. Relevant prior chats are injected as hidden context attachments and shown in the UI as "Grounded with prior chat memory."

This is intentionally local and temporary today: `sessionStorage`, bounded records, no permanent server database. The shape is compatible with a later vector-backed implementation where the same memory records can be embedded, stored per user/file, expired by policy, and retrieved semantically instead of only by keywords.

Production path:

- Prefer MSAL sign-in from the task pane and/or a backend session instead of long-lived local token files.
- Store refresh tokens/secrets server-side using the enterprise vault/session model.
- Start with least-privilege delegated scopes: `openid profile User.Read Files.Read`. Add `offline_access` or broader permissions only through a documented requirement and tenant-consent decision.
- Add broader scopes such as SharePoint site-wide, Mail, Teams, or calendar permissions only when the product explicitly supports those sources and the user can see/approve what context is used.
- Keep source attribution visible in the task pane, e.g. "Using context from FY27 Planning Notes.docx."

## Current limitation

OpenAI-compatible endpoints get native function/tool schemas and a tool-call loop. Anthropic-compatible endpoints now receive Anthropic `tools` schemas and run the same Office/web/M365/generated-file tool executor through `tool_use`/`tool_result` messages. Streaming remains OpenAI-compatible only for now because Anthropic tool streaming and mixed tool/text event handling need a separate production pass.




