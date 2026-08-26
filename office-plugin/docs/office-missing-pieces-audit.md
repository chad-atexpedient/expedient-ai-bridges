# CTRL Office missing-pieces audit

This audit compares the current CTRL add-in capability set against the practical work power users expect from Excel, Word, and PowerPoint. It separates three lanes because Office.js does not expose every ribbon feature equally:

1. **Live tool lane** - direct mutation of the open workbook/document/deck through Office.js.
2. **Generated artifact lane** - server-side Open XML creation or transformation for features Office.js cannot reliably mutate live.
3. **M365/context lane** - delegated Graph/file upload/web/context retrieval that grounds the model before it writes.

## Current shared foundation

Implemented shared capabilities:

- BYOK OpenAI-compatible and Anthropic-compatible provider routing.
- Tool calling with Office, web, image, generated-file, Microsoft 365, and prepared-demo tools.
- Current Office context reading.
- Uploaded Office file and Microsoft 365 file context extraction.
- Generated `.pptx`, `.docx`, and `.xlsx` artifacts.
- Prepared demo showcase via an in-chat feature card by default, with `ctrl_create_demo_showcase` reserved for explicit live Office samples or generated artifact fallback/full-suite mode.
- Skills/agents runtime modules for host-specific behavior.

Missing shared hardening:

| Gap | Why it matters | Best lane | Priority |
| --- | --- | --- | --- |
| Tenant policy for web images | Prevents copyright/security surprises before broad rollout | Server/image import policy | High |
| Production identity/key service | Replaces local BYOK storage with governed user/tenant state | MSAL/backend | High |
| Persisted approved asset library | Lets teams reuse logos/templates safely across sessions | Backend/M365 | Medium |
| Better generated artifact open/import flow | Users should click once and continue in the same Office app | Office/M365/deployment | Medium |
| Runtime API feature telemetry | Helps know which live ribbon-like tools are truly available per tenant/channel | Capability detection | Medium |

## Excel

Current live Excel tool coverage:

- Worksheet create/rename/delete/visibility.
- Range clear/write/formula.
- Structured context imports from Word, PowerPoint, Microsoft 365 files, uploads, web results, or chat into source-attributed Excel tables.
- Charts where runtime APIs expose them.
- Native PivotTables where runtime APIs expose them.
- Pivot-style summary fallback.
- Structured tables.
- Formatting, sorting, filtering, freeze panes.
- Dropdown validation and conditional formatting.
- Images, comments, named ranges, sheet protection.
- Page layout/print setup.
- Formula auditing/risk review sheet generation for volatile formulas, external links, hard-coded constants, long formulas, errors, formula drift, and formula dependency/lineage maps.
- Power Query-style cleanup/combine into new worksheet tables: trim whitespace, normalize repeated spaces, remove blank rows, remove duplicate rows, split a delimited column, apply simple text case cleanup, append two ranges by union/intersection columns, and left-merge lookup columns by key while preserving source ranges.
- Pivot-style chart reports that group/aggregate source data, write a summary table, and create a live chart on a report sheet in one operation.

Current generated Excel coverage:

- Worksheets, typed cells, formulas, styles.
- Tables, validations, conditional formats.
- Charts, embedded images.
- Comments/notes, hyperlinks, named ranges.
- Merged cells, column widths/hidden columns, row heights/hidden rows.
- Frozen panes, view flags, print/page setup, print area, repeating titles.
- Sheet protection and document properties.

Highest-value missing Excel pieces:

| Gap | Examples users expect | Best lane | Notes |
| --- | --- | --- | --- |
| Slicers and timelines | Add slicer to pivot/table, timeline by date | Generated Open XML first, live if Office.js supports enough | High analyst value; current generated workbooks do not create slicer parts. |
| Pivot charts | Pivot chart connected to PivotTable | Generated Open XML first plus live pivot-style report tool | A live `excel_create_pivot_chart_report` bridge now creates grouped summary tables plus live charts in one report sheet. True PivotChart objects connected to PivotTables, slicers/timelines, and generated package-level PivotChart parts remain open. |
| Power Query-style import/transform | Clean CSV/web data, split columns, dedupe, append/merge queries | App/model tool layer + generated workbook | Initial live cleanup and combine tools are implemented for trim/space normalization, blank-row removal, dedupe, split-column, case cleanup, append/union/intersection, and left merge by key into new structured tables. Typed conversion rules, refreshable query definitions, multi-table joins, fuzzy matching, and richer generated-workbook transforms remain open. |
| Advanced formulas/dynamic arrays as guided tools | XLOOKUP, FILTER, LET, LAMBDA, dynamic dashboards | Live formula/write helpers | Could be implemented as higher-level formula tools rather than only raw formula insertion. |
| Scenario/what-if analysis | Goal seek, data tables, solver-like optimization | Generated workbook + model calculation | Solver is not realistic in Office.js; generate formulas/tables and explain assumptions. |
| Workbook auditing | Trace precedents/dependents, formula risk report | Live tool + generated audit sheet | Formula risk audit sheet and dependency-map sheet are implemented; visual trace arrows and richer graph visualization remain. |
| Dashboard layout blocks | KPI cards, variance bridges, waterfall charts | Generated workbook lane | Builds beyond raw charts/tables into polished analyst outputs. |

Recommended next Excel slice: generated workbook **slicers/timeline-style filter controls or pivot-chart packaging**, because those are high-value toolbar features users notice immediately. The first cleanup/combine slice is implemented, so deeper import work should focus on typed conversion, fuzzy matching, refreshable query definitions, and multi-table joins after visible analyst controls unless user testing shows cleanup as the higher pain point.

## Word

Current live Word tool coverage:

- Document audit/review table insertion for structure, readability, accessibility signals, long paragraphs, image alt-text reminders, table-header reminders, raw/generic link cues, and runtime-visible table/image counts.
- Structured context brief insertion from Excel, PowerPoint, Microsoft 365 files, uploads, web results, or chat into a live Word memo/pre-read with headings, source note, summary, key points, and extracted source table where available.
- Text, headings, tables.
- Style application.
- Page and section breaks.
- Primary header/footer.
- Find/replace.
- Comments, images, hyperlinks.
- Selection formatting.
- Guarded content controls where runtime APIs expose them.

Current generated Word coverage:

- Title, TOC field, headings, paragraphs, sections.
- Page layout, margins, orientation, columns.
- Real lists, tables, captions, cross-references.
- Headers/footers, hyperlinks.
- Footnotes/endnotes.
- Comments.
- Tracked-change style insert/delete markup.
- Embedded images.

Highest-value missing Word pieces:

| Gap | Examples users expect | Best lane | Notes |
| --- | --- | --- | --- |
| Source-template preservation | Generate into an uploaded branded/legal `.docx` template | Generated Open XML transform | PowerPoint has template preservation; Word does not yet. This is a major production gap. |
| Rich review workflow | Accept/reject changes, compare documents, consolidate comments | Generated/context + possible live APIs | Live document audit tables, context briefs, and generated redlines exist, but full compare/accept/reject tooling and comment consolidation remain open. |
| Mail merge/document assembly | Generate letters/contracts from Excel/M365 data | Generated Word + Excel/M365 context | High value for sales/legal/admin. |
| Fields beyond TOC/captions | Page numbers, date fields, document properties, REF/PAGEREF polish | Generated Word | Some field support exists; needs broader field catalog and refresh guidance. |
| Style/theme management | Apply document theme, modify styles, enforce brand typography | Generated template lane | Live style application exists, but not style creation/update/theme import. |
| Legal/professional formatting | Clause numbering, defined terms, exhibits, signature blocks | Generated Word | Could become a legal-document skill/tool layer. |
| Accessibility/export checks | Alt text, headings order, table header checks | Context + generated report | Initial live audit table now flags common structure/readability/accessibility signals and reminds users about image alt text and table headers where the runtime can see counts. Full Word Accessibility Checker parity, package-level alt-text extraction, rendered accessibility QA, and export validation remain open. |

Implemented Word slice: **template-preserving DOCX generation** now matches the PowerPoint-upload expectation for the package shell. `word_generate_document_file` accepts an uploaded attachment reference or base64 DOCX, validates required Open XML parts and bounded package size, replaces `word/document.xml`, preserves non-generated template entries such as styles, theme, custom XML, headers/footers, and reusable media, and merges generated relationships/content types. Remaining limitations are intentional: template body content is replaced rather than merged; generated lists may replace template numbering; content-control population, field refresh, malware scanning, and rendered Word layout still require dedicated validation.

Implemented Word slice: **live context briefs** now let Word consume Excel, PowerPoint, Microsoft 365, uploaded, web, or chat context without a raw paste. `word_create_context_brief` inserts a titled brief, source/audience note, executive summary, key-point table, and an extracted source table when the context contains table-like data. This covers common memo/pre-read/summary workflows; deeper narrative synthesis, document comparison, tracked-change acceptance, and mail-merge/document-assembly workflows remain open.

## PowerPoint

Current live PowerPoint tool coverage:

- Deck audit slide insertion for live story/design/readability/accessibility-readiness review using Office.js-visible slides, shapes, and text.
- Slide creation, duplication where supported.
- Text boxes, hyperlink text boxes where supported.
- Tables/native table grids where available.
- Shapes, alignment/distribution/z-order.
- Live editable process, cycle, roadmap, and timeline diagrams using native shapes/text boxes.
- Guarded grouping/ungrouping where runtime APIs expose it.
- Slide background color/image fallback.
- Slide cleanup/delete where supported.
- Image insertion where supported.

Current generated PowerPoint coverage:

- Slide title/body/subtitle.
- Speaker notes through Open XML.
- Background colors.
- Footer/date/classification/slide numbers.
- Shapes.
- Styled table graphic frames.
- Hyperlink text boxes.
- Editable charts with embedded workbook data.
- Embedded images with fit/fill/crop.
- Source `.pptx` template preservation for masters/layouts/themes/fonts/media, with placeholder placement.

Highest-value missing PowerPoint pieces:

| Gap | Examples users expect | Best lane | Notes |
| --- | --- | --- | --- |
| Theme application and brand extraction | Use colors/fonts from uploaded deck or logo automatically | Generated/template + asset analysis | Structured theme extraction and validated fallback application now cover colors/fonts for generated themes, shapes, tables, charts, and chrome; source-template themes remain preserved. Logo ownership/licensing and rendered client QA are still open. |
| Transitions and animations | Fade, morph-like transitions, staged bullet reveals | Generated Open XML | Office.js live support is thin; generated package can add basic transition/animation parts. |
| Speaker notes live mutation | Add notes to current open deck | Live only when proven supported; generated otherwise | Currently intentionally hidden because runtime support failed. |
| Advanced chart types | Waterfall, funnel, map, treemap, richer combo axes | Generated chart XML | Current charts cover common bar/line/pie/area/doughnut/scatter/combo. |
| SmartArt/diagrams | Process, org chart, cycle diagrams | Live/generated shapes first, SmartArt XML later | Initial live editable diagram tool now creates process, cycle, roadmap, and timeline visuals from native slide objects. True SmartArt XML, org-chart-specific layout, and template-aware diagram variants remain open. |
| Slide master/layout editing | Create or modify masters/layouts, not just consume templates | Generated Open XML | Important for production template tooling. |
| Design QA | Alignment, contrast, overflow, consistency checks | Context extraction + visual/render QA | Needs render/screenshot validation path. |

Implemented PowerPoint review slice: **live deck audit slides** now create an editable review slide inside the current deck. The audit uses slide count, runtime-visible object counts, readable text density, blank/dense slide signals, executive-flow cues, and accessibility/design-readiness reminders. This is intentionally not full rendered visual QA; contrast, overflow, exact alignment, theme/media ownership, notes, chart-data, and package-level alt-text checks still require uploaded PPTX extraction, generated artifact QA, or rendered screenshot comparison.

Implemented PowerPoint diagram slice: **live editable diagram generation** now creates process, cycle, roadmap, and timeline visuals in the current deck using Office.js-visible shapes/text boxes. This covers the common "make a workflow/roadmap/cycle" request without sending the user to SmartArt manually or generating a separate download. Remaining limits are true SmartArt XML, org-chart-specific routing, connector geometry beyond simple arrows, and richer template-aware diagram styling.

Implemented PowerPoint slice: **structured theme/brand extraction and fallback style application** now read theme colors, major/minor fonts, layout names/types/placeholders, and bounded media candidates from an uploaded PPTX. The profile is included in uploaded-file context and travels with the source template asset used by generated deck tooling. Fallback generated decks apply validated profile colors/fonts to the generated theme, shapes, tables, charts, and slide chrome; explicit object styles retain precedence and preserved template themes remain unchanged. Remaining work is logo ownership/licensing review and rendered validation across representative Office clients.

## Cross-surface workflows that are still missing

| Workflow | Example | Best lane | Priority |
| --- | --- | --- | --- |
| Word-to-Excel transform | Turn a Word table or narrative into a workbook model | M365/upload/current-context + live Excel import + generated Excel | High; initial live `excel_import_context_table` can convert readable tables, key-value notes, bullets, and plain lines into a structured source-attributed worksheet table. Deeper model-building/formula inference remains open. |
| Excel-to-PowerPoint board pack | Turn workbook tables/charts into branded slides | Generated PPTX + Excel context | High |
| PowerPoint-to-Word brief | Turn a deck into a memo or speaker-notes document | PPTX/current context + live Word brief + generated DOCX | Medium; initial live context-brief tool can consume extracted deck text/tables, but full speaker-note documents and template-preserved generated briefs remain open. |
| Meeting/pre-read assembly | Pull M365 docs, make agenda/deck/report | M365 Graph + generated artifacts | High |
| Multi-artifact project memory | Keep a temporary cache/vector of session artifacts while file is open | Session memory + asset registry | Medium |

## Immediate recommendation

The next three high-impact build slices should be:

1. **Excel slicer/pivot-chart/pivot polish** - closer analyst ribbon parity and better executive workbooks.
2. **Cross-surface transforms** - expand the new Excel context-table import into Excel-to-PowerPoint board packs, PowerPoint-to-Word briefs, and formula/model inference from imported Word/PowerPoint/M365 context.
3. **Rendered Office QA and safety gates** - validate generated/live outputs across representative Office clients, including accessibility, image policy, and package security checks.

These three cover the biggest professional gaps by surface and directly support the product goal: users stay in Office and receive real editable Office outputs rather than instructions.
