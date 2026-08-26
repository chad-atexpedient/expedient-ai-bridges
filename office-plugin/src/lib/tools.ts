import type { OfficeHost, ToolCallRequest, ToolCallResult } from "./types";
import { executeOfficeTool } from "../office/host";
import { primeM365OfficeSso } from "../office/sso";
import { findUploadedAsset, imageAssetFromUrl, listUploadedAssets } from "./uploadRegistry";

export const OFFICE_TOOL_DEFINITIONS = [
  {
    type: "function",
    function: {
      name: "web_search",
      description: "Search the web for current public information. Use this before answering questions that require current facts.",
      parameters: {
        type: "object",
        properties: { query: { type: "string", description: "The web search query." } },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "web_fetch",
      description: "Fetch readable text from a public URL returned by search or provided by the user.",
      parameters: {
        type: "object",
        properties: { url: { type: "string", description: "The public http or https URL to fetch." } },
        required: ["url"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "web_image_search",
      description: "Search the web for public image candidates and return image URLs, page/source URLs, dimensions, and attribution hints. Use this before inserting or embedding topical web images when the user did not provide an image asset.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "The image search query, e.g. dolphins leaping ocean conservation photo." },
          count: { type: "number", description: "Number of image candidates to return. Defaults to 8." },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "web_image_import",
      description: "Fetch, validate, and cache a public image URL as a reusable Office image asset. Use this after web_image_search and before inserting the same web image multiple times or embedding it in generated Office files.",
      parameters: {
        type: "object",
        properties: {
          imageUrl: { type: "string", description: "Public http(s) image URL to validate and cache." },
        },
        required: ["imageUrl"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "m365_try_office_sso",
      description: "Try to authenticate Microsoft 365 Graph access using the user profile already signed into Excel, Word, or PowerPoint through Office SSO. Use this before device login when Microsoft 365 file context is needed.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "m365_auth_status",
      description: "Check whether delegated Microsoft 365 Graph access is ready. If not ready, this returns Microsoft device-login instructions with a user code.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "m365_search_files",
      description: "Search the signed-in user's Microsoft 365 OneDrive/SharePoint files through delegated Microsoft Graph access. Use this to find Word documents, workbooks, decks, or notes that should provide context for the current Office task.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Search terms, such as a document title, project name, or topic." },
          top: { type: "number", description: "Maximum number of files to return. Defaults to 10." },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "m365_read_file",
      description: "Read text context from a Microsoft 365 file returned by m365_search_files. Supports text files and best-effort extraction from Word, PowerPoint, and Excel Open XML files.",
      parameters: {
        type: "object",
        properties: {
          id: { type: "string", description: "Microsoft Graph driveItem id from m365_search_files." },
          maxChars: { type: "number", description: "Maximum characters of extracted text to return. Defaults to 12000." },
        },
        required: ["id"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "office_read_context",
      description: "Read current Excel, Word, or PowerPoint context from the open document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "ctrl_create_demo_showcase",
      description: "Create a prepared CTRL add-in feature demonstration. By default, create the demo live inside the current Excel, Word, or PowerPoint file; use artifact mode or surface=all for safe downloadable Office files.",
      parameters: {
        type: "object",
        properties: {
          surface: { type: "string", enum: ["current", "excel", "word", "powerpoint", "all"], description: "Which demo to create. Defaults to current Office host; use all for a full suite." },
          mode: { type: "string", enum: ["auto", "live", "artifact"], description: "Defaults to auto. Auto creates a live demo in the active Office app when possible, falling back to downloadable artifacts only when live APIs are unavailable. Use artifact for downloadable files." },
          audience: { type: "string", enum: ["executive", "analyst", "sales", "admin", "technical"], description: "Optional audience framing. Defaults to executive." },
          theme: { type: "string", description: "Optional demo theme/topic, such as quarterly business review, sales pipeline, or launch plan." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_get_workbook_overview",
      description: "Read a structural blueprint of the open workbook: sheets, used ranges, header rows, tables, charts, and named ranges. Call this before reading or writing when the workbook layout is unknown.",
      parameters: {
        type: "object",
        properties: {
          includeHeaders: { type: "boolean", description: "Include the first row of each sheet as headers. Defaults to true." },
          maxSheets: { type: "number", description: "Maximum sheets to describe. Defaults to 25." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_read_range",
      description: "Read cell values from the workbook. Use compact for a tab table with addresses, csv for raw data, or detailed to also see formulas and number formats.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string", description: "A1-style range such as B2:F40. Omit to read the current selection." },
          sheetName: { type: "string", description: "Worksheet name. Defaults to the active sheet." },
          mode: { type: "string", enum: ["compact", "csv", "detailed"], description: "Output shape. Defaults to compact." },
          maxRows: { type: "number", description: "Maximum rows to return. Defaults to 200." },
          maxColumns: { type: "number", description: "Maximum columns to return. Defaults to 50." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_search_workbook",
      description: "Find text, values, or formula references across worksheets and return the matching cell addresses.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Text, value, or formula fragment to find." },
          sheetName: { type: "string", description: "Limit the search to one worksheet. Defaults to all worksheets." },
          matchCase: { type: "boolean", description: "Case-sensitive matching. Defaults to false." },
          wholeCell: { type: "boolean", description: "Require the whole cell to equal the query. Defaults to false." },
          searchFormulas: { type: "boolean", description: "Also match formula text. Defaults to true." },
          limit: { type: "number", description: "Maximum matches to return. Defaults to 200." },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_add_worksheet",
      description: "Create and activate a new Excel worksheet.",
      parameters: {
        type: "object",
        properties: { name: { type: "string", description: "Worksheet name, up to 31 characters." } },
        required: ["name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_rename_worksheet",
      description: "Rename an Excel worksheet. Use this for sheet organization instead of asking the user to rename tabs manually.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Current sheet name. Defaults to active sheet." },
          newName: { type: "string", description: "New worksheet name, up to 31 characters." },
        },
        required: ["newName"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_delete_worksheet",
      description: "Delete a named Excel worksheet when the user explicitly asks to remove it.",
      parameters: {
        type: "object",
        properties: { sheetName: { type: "string", description: "Worksheet name to delete." } },
        required: ["sheetName"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_set_worksheet_visibility",
      description: "Show or hide an Excel worksheet for workbook organization.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Worksheet name. Defaults to active sheet." },
          visibility: { type: "string", enum: ["visible", "hidden", "veryHidden"], description: "Desired worksheet visibility." },
        },
        required: ["visibility"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_clear_range",
      description: "Clear values, formats, or all content from an Excel range.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Range address such as A1:D20." },
          applyTo: { type: "string", enum: ["contents", "formats", "all"], description: "What to clear. Defaults to contents." },
        },
        required: ["address"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_write_range",
      description: "Write a two-dimensional array of values into an Excel range, optionally on a named sheet. Writes are refused when the target range already contains data unless overwrite is true, and written values are verified after the write.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Top-left cell address such as A1. Defaults to A1." },
          values: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
          overwrite: { type: "boolean", description: "Set true to replace existing non-empty cells in the target range. Defaults to false, which refuses destructive writes." },
        },
        required: ["values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_import_context_table",
      description: "Turn text context from Word, PowerPoint, Microsoft 365 files, uploads, web results, or chat into a structured Excel worksheet table. Use this after reading external Office context when the user asks to place that information into Excel instead of dumping prose into one cell or one unstructured grid.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Source text to convert into rows. Can be markdown table, CSV/TSV, bullet list, key-value notes, or plain lines." },
          sourceLabel: { type: "string", description: "Short source label such as Planning Notes.docx, PowerPoint audit, web search, or Current Word selection." },
          outputSheetName: { type: "string", description: "Worksheet name for the imported table. Defaults to Imported Context." },
          outputAddress: { type: "string", description: "Top-left output cell. Defaults to A1." },
          mode: { type: "string", enum: ["auto", "table", "keyValue", "bullets", "lines"], description: "Parsing mode. Defaults to auto." },
          includeSourceColumn: { type: "boolean", description: "Add a Source column to preserve provenance. Defaults to true." },
          clean: { type: "boolean", description: "Trim whitespace, normalize repeated spaces, and remove blanks. Defaults to true." },
          removeDuplicateRows: { type: "boolean", description: "Remove duplicate imported rows. Defaults to false." },
          preserveFormulas: { type: "boolean", description: "Allow imported cells beginning with =, +, -, or @ to become formulas. Defaults to false for safety." },
          createTable: { type: "boolean", description: "Create an Excel structured table when possible. Defaults to true." },
          tableName: { type: "string", description: "Optional structured table name." },
          maxRows: { type: "number", description: "Maximum parsed rows to import, capped internally. Defaults to 500." },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_set_formula",
      description: "Set a formula in a single Excel cell.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Cell address such as C2." },
          formula: { type: "string", description: "Excel formula, beginning with =." },
        },
        required: ["address", "formula"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_clean_transform_range",
      description: "Clean and transform an Excel source range into a new structured output table. Use this for Power Query-style cleanup tasks such as trimming whitespace, normalizing repeated spaces, removing blank rows, deduplicating records, splitting a column by delimiter, and applying simple case cleanup without altering the source range.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional source worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Source range including headers, such as A1:F500. Defaults to the worksheet used range when omitted." },
          outputSheetName: { type: "string", description: "Worksheet to write cleaned output to. Defaults to Cleaned Data." },
          outputAddress: { type: "string", description: "Top-left output cell. Defaults to A1." },
          trimWhitespace: { type: "boolean", description: "Trim leading/trailing whitespace from text cells. Defaults to true." },
          normalizeSpaces: { type: "boolean", description: "Collapse repeated whitespace inside text cells. Defaults to true." },
          removeBlankRows: { type: "boolean", description: "Remove blank data rows. Defaults to true." },
          removeDuplicateRows: { type: "boolean", description: "Remove duplicate data rows after cleanup. Defaults to false." },
          splitColumn: { type: ["string", "number"], description: "Optional header name or zero-based index of a column to split by delimiter." },
          delimiter: { type: "string", description: "Delimiter for splitColumn. Defaults to comma." },
          hasHeaders: { type: "boolean", description: "Whether the first row contains headers. Defaults to true." },
          caseMode: { type: "string", enum: ["none", "upper", "lower", "title"], description: "Optional text case cleanup. Defaults to none." },
          createTable: { type: "boolean", description: "Create an Excel structured table around the cleaned output. Defaults to true." },
          tableName: { type: "string", description: "Optional structured table name." },
          style: { type: "string", description: "Optional Excel table style, e.g. TableStyleMedium2." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_combine_ranges",
      description: "Append or merge two Excel ranges into a new structured output table. Use this for Power Query-style combine workflows: stack rows from two sheets/tables with matching or unioned columns, or left-merge lookup columns from a secondary range by key without changing the source ranges.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name used for both ranges when primarySheetName/secondarySheetName are omitted." },
          primarySheetName: { type: "string", description: "Optional primary/source worksheet name. Defaults to active sheet." },
          primaryAddress: { type: "string", description: "Primary range including headers, such as A1:D200." },
          secondarySheetName: { type: "string", description: "Optional secondary/lookup worksheet name. Defaults to active sheet." },
          secondaryAddress: { type: "string", description: "Secondary range including headers, such as F1:H200." },
          mode: { type: "string", enum: ["append", "merge"], description: "append stacks both ranges. merge performs a left lookup-style merge from secondary into primary. Defaults to append." },
          matchColumns: { type: "string", enum: ["union", "intersection"], description: "Append behavior: union keeps every column from both ranges; intersection keeps matching columns only. Defaults to union." },
          primaryKey: { type: ["string", "number"], description: "Primary key header/index for merge mode." },
          secondaryKey: { type: ["string", "number"], description: "Secondary key header/index for merge mode. Defaults to primaryKey." },
          conflictSuffix: { type: "string", description: "Suffix for secondary columns that conflict with primary column names in merge mode. Defaults to ' from lookup'." },
          includeSourceColumn: { type: "boolean", description: "For append mode, add a Source column identifying which range each row came from. Defaults to false." },
          primaryLabel: { type: "string", description: "Source label for primary rows when includeSourceColumn is true. Defaults to primary sheet name." },
          secondaryLabel: { type: "string", description: "Source label for secondary rows when includeSourceColumn is true. Defaults to secondary sheet name." },
          outputSheetName: { type: "string", description: "Worksheet to write combined output to. Defaults to Appended Data or Merged Data." },
          outputAddress: { type: "string", description: "Top-left output cell. Defaults to A1." },
          trimWhitespace: { type: "boolean", description: "Trim leading/trailing whitespace from text cells before combining. Defaults to true." },
          normalizeSpaces: { type: "boolean", description: "Collapse repeated whitespace inside text cells before combining. Defaults to true." },
          removeBlankRows: { type: "boolean", description: "Remove blank data rows before combining. Defaults to true." },
          removeDuplicateRows: { type: "boolean", description: "Remove duplicate rows within each source before combining. Defaults to false." },
          hasHeaders: { type: "boolean", description: "Whether each range has a header row. Defaults to true." },
          caseMode: { type: "string", enum: ["none", "upper", "lower", "title"], description: "Optional text case cleanup before combining. Defaults to none." },
          createTable: { type: "boolean", description: "Create an Excel structured table around the combined output. Defaults to true." },
          tableName: { type: "string", description: "Optional structured table name." },
          style: { type: "string", description: "Optional Excel table style, e.g. TableStyleMedium4." },
        },
        required: ["primaryAddress", "secondaryAddress"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_audit_formulas",
      description: "Audit formulas in an Excel worksheet or range and create a professional review sheet with risks such as errors, volatile functions, external links, hard-coded constants, long formulas, and inconsistent formulas down a column. Use this for Formula Auditing-style workbook review instead of only explaining what to check manually.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet to audit. Defaults to active sheet." },
          address: { type: "string", description: "Optional range to audit, such as A1:K200. Defaults to the worksheet used range." },
          outputSheetName: { type: "string", description: "Worksheet for the audit report. Defaults to Formula Audit." },
          includeLowRisk: { type: "boolean", description: "Include low-severity informational findings. Defaults to false." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_map_formula_dependencies",
      description: "Map formula precedents/dependencies in an Excel worksheet or range and create a professional dependency review sheet. Use this for Trace Precedents/Dependents-style workbook review, model lineage, and finance-model documentation instead of asking the user to click formula arrows manually.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet to map. Defaults to active sheet." },
          address: { type: "string", description: "Optional range to map, such as A1:K200. Defaults to the worksheet used range." },
          outputSheetName: { type: "string", description: "Worksheet for the dependency map. Defaults to Dependency Map." },
          includeSummary: { type: "boolean", description: "Include dependent summary rows after the precedent map. Defaults to true." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_create_chart",
      description: "Create an Excel chart object from a range on the active or named worksheet.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Source range address such as A5:C17." },
          chartType: { type: "string", enum: ["line", "lineMarkers", "columnClustered", "barClustered", "pie", "scatter", "area"], description: "Chart type. Use lineMarkers for a line chart with markers." },
          title: { type: "string", description: "Optional chart title." },
          startCell: { type: "string", description: "Optional chart anchor cell such as E2. Defaults to E2." },
          endCell: { type: "string", description: "Optional opposite chart anchor cell such as M20. Defaults to M20." },
          top: { type: "number", description: "Optional chart top position in points." },
          left: { type: "number", description: "Optional chart left position in points." },
          width: { type: "number", description: "Optional chart width in points." },
          height: { type: "number", description: "Optional chart height in points." },
        },
        required: ["sourceAddress", "chartType"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_create_summary_table",
      description: "Create a pivot-style Excel summary table by grouping a source range by one column and aggregating another column with sum, average, count, min, or max. Use this for quick professional analysis when a true PivotTable API is unavailable or unnecessary.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional source worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Source range including headers, such as A1:D200." },
          groupBy: { type: ["string", "number"], description: "Header name or zero-based column index to group by." },
          valueColumn: { type: ["string", "number"], description: "Header name or zero-based column index to aggregate." },
          operation: { type: "string", enum: ["sum", "average", "count", "min", "max"], description: "Aggregation operation. Defaults to sum." },
          outputSheetName: { type: "string", description: "Worksheet to write the summary table to. Defaults to AI Summary." },
          outputAddress: { type: "string", description: "Top-left output cell. Defaults to A1." },
          tableName: { type: "string", description: "Optional structured table name." },
          style: { type: "string", description: "Optional Excel table style, e.g. TableStyleMedium4." },
        },
        required: ["sourceAddress", "groupBy", "valueColumn"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_create_pivot_chart_report",
      description: "Create a pivot-style Excel report sheet with a grouped summary table and chart from a source range. Use this when the user asks for a PivotChart, charted summary, grouped chart, executive analyst view, or quick visual analysis and native PivotChart APIs are unavailable or unnecessary.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional source worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Source range including headers, such as A1:F500." },
          groupBy: { type: ["string", "number"], description: "Header name or zero-based column index to group by." },
          valueColumn: { type: ["string", "number"], description: "Header name or zero-based column index to aggregate." },
          operation: { type: "string", enum: ["sum", "average", "count", "min", "max"], description: "Aggregation operation. Defaults to sum." },
          outputSheetName: { type: "string", description: "Worksheet to write the report to. Defaults to Pivot Chart Report." },
          outputAddress: { type: "string", description: "Top-left output cell for the summary table. Defaults to A1." },
          tableName: { type: "string", description: "Optional structured table name for the summary data." },
          style: { type: "string", description: "Optional Excel table style, e.g. TableStyleMedium4." },
          chartType: { type: "string", enum: ["line", "lineMarkers", "columnClustered", "barClustered", "pie", "scatter", "area"], description: "Chart type. Defaults to columnClustered." },
          title: { type: "string", description: "Optional chart title." },
          chartStartCell: { type: "string", description: "Chart anchor cell such as E2. Defaults to E2." },
          chartEndCell: { type: "string", description: "Opposite chart anchor cell such as M20. Defaults to M20." },
          top: { type: "number", description: "Optional top N groups to keep after sorting." },
          sortBy: { type: "string", enum: ["label", "value", "valueDesc", "valueAsc"], description: "Sort groups by label, descending value, or ascending value. Defaults to label." },
        },
        required: ["sourceAddress", "groupBy", "valueColumn"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_create_pivottable",
      description: "Create a real Excel PivotTable from a source range when the Excel runtime exposes PivotTable APIs. Use this for professional analyst summaries by rows, columns, filters, and aggregated value fields; fall back to excel_create_summary_table or excel_generate_workbook_file if the runtime cannot create native PivotTables.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional source worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Source range including headers, such as A1:F500." },
          destinationSheetName: { type: "string", description: "Worksheet for the PivotTable. Defaults to Pivot Summary." },
          destinationAddress: { type: "string", description: "Top-left cell for the PivotTable. Defaults to A3." },
          name: { type: "string", description: "Optional PivotTable name." },
          rows: { type: "array", items: { type: ["string", "number"] }, description: "Header names or zero-based source column indexes to add as row fields." },
          columns: { type: "array", items: { type: ["string", "number"] }, description: "Header names or zero-based source column indexes to add as column fields." },
          filters: { type: "array", items: { type: ["string", "number"] }, description: "Header names or zero-based source column indexes to add as report filter fields." },
          values: { type: "array", items: { type: "object", properties: { field: { type: ["string", "number"] }, summarizeBy: { type: "string", enum: ["sum", "count", "average", "max", "min"] }, name: { type: "string" } }, required: ["field"] }, description: "Value fields to aggregate. Defaults to sum unless summarizeBy is supplied." },
          layout: { type: "string", enum: ["compact", "outline", "tabular"], description: "Optional report layout when supported." },
        },
        required: ["sourceAddress"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_create_table",
      description: "Convert a worksheet range into a structured Excel table with optional name and style.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          sourceAddress: { type: "string", description: "Range address such as A1:D20." },
          name: { type: "string", description: "Optional Excel table name." },
          hasHeaders: { type: "boolean", description: "Whether the first row contains headers. Defaults to true." },
          style: { type: "string", description: "Optional table style, e.g. TableStyleMedium2." },
        },
        required: ["sourceAddress"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_format_range",
      description: "Apply professional formatting to an Excel range: number format, bold, colors, font size, alignment, wrap, and autofit.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Range address such as A1:D20." },
          numberFormat: { type: "string", description: "Optional Excel number format such as $#,##0.00 or 0.0%." },
          bold: { type: "boolean" },
          fillColor: { type: "string", description: "Hex color such as #1f4e79." },
          fontColor: { type: "string", description: "Hex color such as #ffffff." },
          fontSize: { type: "number" },
          horizontalAlignment: { type: "string", enum: ["left", "center", "right", "justify"] },
          wrapText: { type: "boolean" },
          autofit: { type: "boolean", description: "Autofit rows and columns. Defaults to true." },
        },
        required: ["address"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_sort_range",
      description: "Sort an Excel range by a zero-based column index within that range.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Range address such as A1:D20." },
          keyColumnIndex: { type: "number", description: "Zero-based column index within the range to sort by." },
          ascending: { type: "boolean", description: "Defaults to true." },
          hasHeaders: { type: "boolean", description: "Defaults to true." },
        },
        required: ["address", "keyColumnIndex"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_filter_range",
      description: "Apply or clear a filter on an Excel table or range by column values/criteria.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          tableName: { type: "string", description: "Optional table name. If omitted, address is used." },
          address: { type: "string", description: "Range address when filtering a plain range." },
          columnIndex: { type: "number", description: "Zero-based column index to filter." },
          values: { type: "array", items: { type: ["string", "number", "boolean"] }, description: "Allowed values." },
          clear: { type: "boolean", description: "Clear filters instead of applying criteria." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_freeze_panes",
      description: "Freeze worksheet panes by rows, columns, or at a specific cell.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          rows: { type: "number", description: "Number of top rows to freeze." },
          columns: { type: "number", description: "Number of left columns to freeze." },
          cell: { type: "string", description: "Optional split cell such as B2." },
          clear: { type: "boolean", description: "Unfreeze panes." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_add_data_validation",
      description: "Add dropdown list data validation to an Excel range.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Range address such as D2:D100." },
          listValues: { type: "array", items: { type: "string" }, description: "Dropdown values." },
          promptTitle: { type: "string" },
          promptMessage: { type: "string" },
          errorTitle: { type: "string" },
          errorMessage: { type: "string" },
        },
        required: ["address", "listValues"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_apply_conditional_format",
      description: "Apply a simple conditional format to an Excel range using greaterThan, lessThan, equalTo, containsText, or colorScale.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          address: { type: "string", description: "Range address such as C2:C100." },
          ruleType: { type: "string", enum: ["greaterThan", "lessThan", "equalTo", "containsText", "colorScale"] },
          value: { type: ["string", "number"], description: "Comparison value for non-colorScale rules." },
          fillColor: { type: "string", description: "Hex fill color for matching cells." },
          fontColor: { type: "string", description: "Hex font color for matching cells." },
        },
        required: ["address", "ruleType"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_insert_image",
      description: "Insert an uploaded or web image into the active or named Excel worksheet as a real image object. Use assetId/assetName for attached images or imageUrl for public web images.",
      parameters: {
        type: "object",
        properties: {
          sheetName: { type: "string", description: "Optional worksheet name. Defaults to active sheet." },
          assetId: { type: "string", description: "Uploaded image asset id." },
          assetName: { type: "string", description: "Uploaded image filename." },
          imageUrl: { type: "string", description: "Public http(s) image URL to fetch and insert." },
          startCell: { type: "string", description: "Optional anchor cell such as E2." },
          left: { type: "number", description: "Optional left position in points." },
          top: { type: "number", description: "Optional top position in points." },
          width: { type: "number", description: "Optional width in points." },
          height: { type: "number", description: "Optional height in points." },
          altText: { type: "string", description: "Optional accessibility description." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_add_comment",
      description: "Add a real Excel comment/note to a cell or range, useful for review workflows instead of placing comment text in nearby cells.",
      parameters: { type: "object", properties: { sheetName: { type: "string" }, address: { type: "string", description: "Cell or range address such as B4." }, text: { type: "string", description: "Comment text." } }, required: ["address", "text"] },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_set_named_range",
      description: "Create or update a workbook named range so formulas and downstream automation can reference a professional semantic name.",
      parameters: { type: "object", properties: { name: { type: "string" }, sheetName: { type: "string" }, address: { type: "string" } }, required: ["name", "address"] },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_protect_sheet",
      description: "Protect or unprotect a worksheet with optional common permissions. Use this for professional workbook hardening when the user asks to lock or protect a sheet.",
      parameters: { type: "object", properties: { sheetName: { type: "string" }, protect: { type: "boolean" }, password: { type: "string" }, allowFormatCells: { type: "boolean" }, allowSort: { type: "boolean" }, allowAutoFilter: { type: "boolean" } } },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_set_page_layout",
      description: "Set professional worksheet print/page setup: orientation, paper size, margins, print area, repeating title rows/columns, fit-to-page scaling, gridlines/headings, and page centering instead of asking the user to open Page Layout manually.",
      parameters: { type: "object", properties: { sheetName: { type: "string" }, orientation: { type: "string", enum: ["portrait", "landscape"] }, paperSize: { type: "string", enum: ["letter", "legal", "a4"] }, topMargin: { type: "number" }, bottomMargin: { type: "number" }, leftMargin: { type: "number" }, rightMargin: { type: "number" }, printArea: { type: "string", description: "Range to print, such as A1:G45. Empty string clears when supported." }, repeatRows: { type: "string", description: "Rows to repeat at top, such as 1:2 or $1:$2." }, repeatColumns: { type: "string", description: "Columns to repeat at left, such as A:B or $A:$B." }, fitToPagesWide: { type: "number", description: "Fit printed output to this many pages wide." }, fitToPagesTall: { type: "number", description: "Fit printed output to this many pages tall." }, scale: { type: "number", description: "Print scale percentage, typically 10-400. Use only when not fitting to pages." }, showGridlines: { type: "boolean" }, showHeadings: { type: "boolean" }, centerHorizontally: { type: "boolean" }, centerVertically: { type: "boolean" }, blackAndWhite: { type: "boolean" }, draftMode: { type: "boolean" } } },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_text",
      description: "Insert text into the Word document at the selection, beginning, or end.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string" },
          location: { type: "string", enum: ["replace", "start", "end"] },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_audit_document",
      description: "Audit the current Word document for professional structure, readability, and accessibility signals, then insert a real Word review table at the end of the document. Use this for document QA, accessibility review, executive polish checks, and professional review workflows instead of only giving generic advice.",
      parameters: {
        type: "object",
        properties: {
          includeHeading: { type: "boolean", description: "Insert a heading before the audit table. Defaults to true." },
          style: { type: "string", description: "Optional Word table style name. Defaults to Grid Table 4 - Accent 1." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_create_context_brief",
      description: "Create a structured Word brief in the current document from context read from Excel, PowerPoint, Microsoft 365 files, uploads, web results, or chat. Use this when the user asks to turn cross-surface context into a memo, brief, summary, pre-read, or narrative Word artifact instead of pasting raw context.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Source context to brief. Can contain prose, bullets, key-value notes, markdown/CSV/TSV tables, or extracted Office text." },
          title: { type: "string", description: "Brief title. Defaults to Context Brief." },
          sourceLabel: { type: "string", description: "Short source label such as Sales Pipeline.xlsx, Board Deck.pptx, M365 search result, or web research." },
          audience: { type: "string", enum: ["executive", "sales", "technical", "legal", "professional"], description: "Audience framing. Defaults to professional." },
          location: { type: "string", enum: ["replace", "start", "end"], description: "Where to insert the brief. Defaults to end." },
          includeSourceTable: { type: "boolean", description: "Include an extracted source table when one can be detected. Defaults to true." },
          maxPoints: { type: "number", description: "Maximum key points. Defaults to 8." },
          maxTableRows: { type: "number", description: "Maximum source table rows to include. Defaults to 20." },
          style: { type: "string", description: "Optional Word table style for inserted tables. Defaults to Grid Table 4 - Accent 1." },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_heading",
      description: "Insert a styled Word heading at the selection, beginning, or end of the document.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string" },
          level: { type: "number", description: "Heading level 1-6. Defaults to 1." },
          location: { type: "string", enum: ["replace", "start", "end"] },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_table",
      description: "Insert a Word table from a two-dimensional array of values at the selection, beginning, or end.",
      parameters: {
        type: "object",
        properties: {
          values: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
          location: { type: "string", enum: ["replace", "start", "end"] },
          style: { type: "string", description: "Optional Word table style name." },
        },
        required: ["values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_apply_style",
      description: "Apply a Word style to the current selection or whole body.",
      parameters: {
        type: "object",
        properties: {
          style: { type: "string", description: "Style name such as Normal, Title, Heading 1, Quote." },
          target: { type: "string", enum: ["selection", "body"] },
        },
        required: ["style"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_page_break",
      description: "Insert a page break at the current selection, beginning, or end of the document.",
      parameters: {
        type: "object",
        properties: {
          location: { type: "string", enum: ["replace", "start", "end"] },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_section_break",
      description: "Insert a Word section break, useful for professional reports with different layouts, headers, or pagination.",
      parameters: {
        type: "object",
        properties: {
          breakType: { type: "string", enum: ["nextPage", "continuous"], description: "Section break type. Defaults to nextPage." },
          location: { type: "string", enum: ["replace", "start", "end"], description: "Where to insert. Defaults to current selection." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_set_header_footer",
      description: "Set the primary Word document header and/or footer text for report-style documents.",
      parameters: {
        type: "object",
        properties: {
          header: { type: "string", description: "Header text to set. Omit to leave unchanged." },
          footer: { type: "string", description: "Footer text to set. Omit to leave unchanged." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_find_replace",
      description: "Find and replace text in the current Word document body.",
      parameters: {
        type: "object",
        properties: {
          find: { type: "string" },
          replace: { type: "string" },
          matchCase: { type: "boolean" },
          matchWholeWord: { type: "boolean" },
        },
        required: ["find", "replace"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_comment",
      description: "Insert a Word comment on the current selection when the runtime supports comments.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_image",
      description: "Insert an uploaded or web image into Word at the current selection, beginning, or end. Use assetId/assetName for attached images or imageUrl for public web images.",
      parameters: {
        type: "object",
        properties: {
          assetId: { type: "string", description: "Uploaded image asset id." },
          assetName: { type: "string", description: "Uploaded image filename." },
          imageUrl: { type: "string", description: "Public http(s) image URL to fetch and insert." },
          location: { type: "string", enum: ["replace", "start", "end"] },
          width: { type: "number", description: "Optional width in points." },
          height: { type: "number", description: "Optional height in points." },
          altText: { type: "string", description: "Optional accessibility description." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_hyperlink",
      description: "Insert a real Word hyperlink at the selection, beginning, or end using HTML insertion where supported.",
      parameters: { type: "object", properties: { text: { type: "string" }, url: { type: "string" }, location: { type: "string", enum: ["replace", "start", "end"] } }, required: ["text", "url"] },
    },
  },
  {
    type: "function",
    function: {
      name: "word_format_selection",
      description: "Apply common font/highlight formatting to the current Word selection or whole body, like bold, italic, font size, colors, and highlight.",
      parameters: { type: "object", properties: { target: { type: "string", enum: ["selection", "body"] }, bold: { type: "boolean" }, italic: { type: "boolean" }, underline: { type: "boolean" }, fontSize: { type: "number" }, fontColor: { type: "string" }, highlightColor: { type: "string" } } },
    },
  },
  {
    type: "function",
    function: {
      name: "word_insert_content_control",
      description: "Insert or wrap the current Word selection in a content control for professional templates, reusable fields, contracts, forms, and guided document assembly. Supports richText/plainText/checkBox/dropdown/date intents where the Word runtime exposes them.",
      parameters: { type: "object", properties: { title: { type: "string", description: "Visible content-control title." }, tag: { type: "string", description: "Machine-readable tag/key for automation." }, placeholderText: { type: "string", description: "Placeholder or initial text." }, type: { type: "string", enum: ["richText", "plainText", "checkBox", "dropdown", "date"], description: "Desired control type. Defaults to richText." }, location: { type: "string", enum: ["replace", "start", "end", "selection"], description: "Where to insert if not wrapping selection. Defaults to selection/replace." }, options: { type: "array", items: { type: "string" }, description: "Dropdown/list choices when type is dropdown." }, cannotDelete: { type: "boolean" }, cannotEdit: { type: "boolean" } }, required: ["title"] },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_create_slides",
      description: "Create one or more PowerPoint slides with title and body text boxes.",
      parameters: {
        type: "object",
        properties: {
          slides: {
            type: "array",
            items: {
              type: "object",
              properties: { title: { type: "string" }, body: { type: "string" } },
              required: ["title", "body"],
            },
          },
        },
        required: ["slides"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_audit_deck",
      description: "Review the active PowerPoint deck using live slide/object/text context and add a concrete audit slide with findings and recommendations. Use this for deck QA, design review, executive polish, accessibility/design-readiness checks, and slide-story review instead of only giving generic presentation advice.",
      parameters: {
        type: "object",
        properties: {
          audience: { type: "string", description: "Optional audience framing such as executive, sales, board, technical, or training. Defaults to executive." },
          title: { type: "string", description: "Optional audit slide title. Defaults to CTRL PowerPoint Deck Audit." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_add_textbox",
      description: "Add a positioned text box to a PowerPoint slide with optional font styling.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide, creating one if needed." },
          text: { type: "string" },
          left: { type: "number" },
          top: { type: "number" },
          width: { type: "number" },
          height: { type: "number" },
          fontSize: { type: "number" },
          bold: { type: "boolean" },
          fontColor: { type: "string", description: "Hex color such as #1f4e79." },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_add_table",
      description: "Add a table-like grid of content to a PowerPoint slide. Uses native tables when available, otherwise creates positioned text boxes as a reliable visual table.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide, creating one if needed." },
          values: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
          left: { type: "number" },
          top: { type: "number" },
          width: { type: "number" },
          height: { type: "number" },
          headerFillColor: { type: "string", description: "Optional header fill color." },
          fontSize: { type: "number" },
        },
        required: ["values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_add_shape",
      description: "Add a positioned PowerPoint shape such as rectangle, oval, triangle, line, or roundedRectangle, with optional text and styling.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide, creating one if needed." },
          shapeType: { type: "string", enum: ["rectangle", "roundedRectangle", "oval", "triangle", "line"] },
          text: { type: "string" },
          left: { type: "number" },
          top: { type: "number" },
          width: { type: "number" },
          height: { type: "number" },
          fillColor: { type: "string" },
          lineColor: { type: "string" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_create_diagram",
      description: "Create a live editable PowerPoint diagram on a slide using native shapes/text boxes. Use this for process flows, cycles, roadmaps, timelines, workflows, and simple visual frameworks instead of manual SmartArt instructions or downloadable demo decks.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Prefer slideNumber for user-facing references." },
          slideNumber: { type: "number", description: "Human 1-based slide number." },
          slidePosition: { type: "string", enum: ["last"], description: "Use last to place the diagram on the last slide." },
          diagramType: { type: "string", enum: ["process", "cycle", "roadmap", "timeline"], description: "Diagram layout. Defaults to process." },
          title: { type: "string", description: "Optional title above the diagram." },
          steps: {
            type: "array",
            items: {
              oneOf: [
                { type: "string" },
                { type: "object", properties: { label: { type: "string" }, detail: { type: "string" }, owner: { type: "string" }, status: { type: "string" } }, required: ["label"] },
              ],
            },
            description: "Diagram steps/cards. Keep labels concise; use detail for secondary text.",
          },
          left: { type: "number" },
          top: { type: "number" },
          width: { type: "number" },
          height: { type: "number" },
          fillColor: { type: "string", description: "Card fill color, such as #dbeafe." },
          accentColor: { type: "string", description: "Connector/title/border accent color, such as #2563eb." },
          fontColor: { type: "string", description: "Text color, such as #111827." },
          numbered: { type: "boolean", description: "Whether to prefix step labels with numbers. Defaults to true except cycle diagrams." },
        },
        required: ["steps"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_arrange_shapes",
      description: "Arrange existing PowerPoint shapes on a slide: align left/center/right/top/middle/bottom, distribute horizontally/vertically, or send/bring shapes forward/back when supported. Uses 1-based shapeNumbers from current slide order; omit shapeNumbers to arrange all shapes on the slide.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Prefer slideNumber for user-facing references." },
          slideNumber: { type: "number", description: "Human 1-based slide number." },
          shapeNumbers: { type: "array", items: { type: "number" }, description: "Optional 1-based shape numbers to arrange. Defaults to all shapes on the slide." },
          action: { type: "string", enum: ["alignLeft", "alignCenter", "alignRight", "alignTop", "alignMiddle", "alignBottom", "distributeHorizontal", "distributeVertical"], description: "Alignment or distribution action." },
          order: { type: "string", enum: ["bringToFront", "sendToBack", "bringForward", "sendBackward"], description: "Optional z-order action when runtime supports it." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_group_shapes",
      description: "Group or ungroup existing PowerPoint shapes when the installed PowerPoint runtime exposes grouping APIs. Uses 1-based shapeNumbers from current slide order; group requires at least two shapes.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Prefer slideNumber for user-facing references." },
          slideNumber: { type: "number", description: "Human 1-based slide number." },
          shapeNumbers: { type: "array", items: { type: "number" }, description: "1-based shape numbers to group or ungroup. Omit only for ungrouping all group-like shapes on the slide." },
          action: { type: "string", enum: ["group", "ungroup"], description: "Whether to group selected shapes or ungroup selected group shapes. Defaults to group." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_delete_slide",
      description: "Delete a PowerPoint slide by zero-based slide index when the runtime supports deleting slides.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index to delete." },
        },
        required: ["slideIndex"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_clear_slide",
      description: "Remove all shapes from a PowerPoint slide when the runtime supports shape deletion.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_set_slide_background",
      description: "Set a PowerPoint slide background color or background image. Use this for professional slide design instead of leaving generic blank slides.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide, creating one if needed." },
          color: { type: "string", description: "Hex background color such as #0f172a." },
          assetId: { type: "string", description: "Uploaded image asset id for a full-slide background image." },
          assetName: { type: "string", description: "Uploaded image filename for a full-slide background image." },
          imageUrl: { type: "string", description: "Public http(s) image URL for a full-slide background image." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_add_speaker_notes",
      description: "Add speaker notes to a PowerPoint slide when the runtime exposes notes APIs. Returns a clear limitation otherwise.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to first slide." },
          notes: { type: "string" },
        },
        required: ["notes"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_insert_image",
      description: "Insert an uploaded or web image into a PowerPoint slide as a real image object. Use assetId/assetName for attached images or imageUrl for public web images.",
      parameters: {
        type: "object",
        properties: {
          slideIndex: { type: "number", description: "Zero-based slide index. Defaults to the first slide, creating one if needed." },
          assetId: { type: "string", description: "Uploaded image asset id." },
          assetName: { type: "string", description: "Uploaded image filename." },
          imageUrl: { type: "string", description: "Public http(s) image URL to fetch and insert." },
          left: { type: "number", description: "Optional left position in points." },
          top: { type: "number", description: "Optional top position in points." },
          width: { type: "number", description: "Optional width in points." },
          height: { type: "number", description: "Optional height in points." },
          altText: { type: "string", description: "Optional accessibility description." },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_duplicate_slide",
      description: "Duplicate an existing PowerPoint slide when the runtime exposes slide duplication. Use this for template-like deck building instead of recreating slide layouts manually.",
      parameters: { type: "object", properties: { slideIndex: { type: "number", description: "Zero-based slide index to duplicate. Defaults to first slide." } } },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_add_hyperlink_textbox",
      description: "Add a PowerPoint text box intended as a real hyperlink when the runtime exposes hyperlink APIs, falling back with a clear limitation if unsupported.",
      parameters: { type: "object", properties: { slideIndex: { type: "number" }, text: { type: "string" }, url: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" } }, required: ["text", "url"] },
    },
  },
  {
    type: "function",
    function: {
      name: "word_generate_document_file",
      description: "Generate a downloadable .docx file server-side when a standalone Word artifact is better than live document mutation. Supports title, table of contents field, page layout, section-specific layout, columns, sections, paragraphs, real bulleted/numbered lists, tables, real SEQ-field captions/bookmarks, REF-field cross-references, headers, footers, footnotes, endnotes, hyperlink relationships, comment parts, tracked-change revision markup, real embedded images, and bounded template-shell preservation from an uploaded/source DOCX.",
      parameters: {
        type: "object",
        properties: {
          fileName: { type: "string" },
          title: { type: "string" },
          body: { type: "string" },
          templateAssetId: { type: "string", description: "Optional uploaded DOCX template asset id. The template package shell is preserved while CTRL replaces the generated document body." },
          templateAssetName: { type: "string", description: "Optional uploaded DOCX template filename to use as the source template." },
          templateBase64: { type: "string", description: "Optional base64 or data URL of a bounded .docx template package. Prefer templateAssetId/templateAssetName for attached files." },
          tableOfContents: { type: "boolean", description: "Include a Word TOC field that can be refreshed in Word." },
          page: { type: "object", properties: { orientation: { type: "string", enum: ["portrait", "landscape"] }, margins: { type: "object", properties: { top: { type: "number" }, right: { type: "number" }, bottom: { type: "number" }, left: { type: "number" } } }, columns: { type: "object", properties: { count: { type: "number" }, space: { type: "number", description: "Column gap in twips. 720 is 0.5 inch." }, separator: { type: "boolean" } } } } },
          columns: { type: "object", properties: { count: { type: "number", description: "Number of Word text columns, 1-8." }, space: { type: "number", description: "Column gap in twips. 720 is 0.5 inch." }, separator: { type: "boolean", description: "Show a vertical separator line between columns." } }, description: "Document-level Word columns for generated DOCX output." },
          header: { type: "string" },
          footer: { type: "string" },
          images: {
            type: "array",
            items: { type: "object", properties: { name: { type: "string" }, type: { type: "string" }, assetId: { type: "string" }, assetName: { type: "string" }, imageUrl: { type: "string" }, base64: { type: "string" }, width: { type: "number" }, height: { type: "number" }, altText: { type: "string" }, caption: { type: "object", properties: { label: { type: "string" }, text: { type: "string" }, bookmark: { type: "string" } } } } },
          },
          comments: { type: "array", items: { type: "object", properties: { text: { type: "string" }, author: { type: "string" } }, required: ["text"] } },
          captions: { type: "array", items: { type: "object", properties: { label: { type: "string", description: "Caption label such as Figure, Table, or Exhibit." }, text: { type: "string" }, bookmark: { type: "string", description: "Bookmark name used by crossReferences." } } }, description: "Standalone Word captions using SEQ fields and bookmarks." },
          crossReferences: { type: "array", items: { type: "object", properties: { bookmark: { type: "string" }, text: { type: "string", description: "Prefix text, such as See or Refer to." }, fallback: { type: "string", description: "Visible fallback until Word fields refresh." } }, required: ["bookmark"] }, description: "Word REF fields pointing to generated caption bookmarks." },
          revisions: { type: "array", items: { type: "object", properties: { type: { type: "string", enum: ["insert", "delete"] }, text: { type: "string" }, author: { type: "string" }, date: { type: "string" } }, required: ["type", "text"] }, description: "Top-level tracked-change style insertions/deletions using WordprocessingML w:ins/w:del markup." },
          links: { type: "array", items: { type: "object", properties: { text: { type: "string" }, label: { type: "string" }, url: { type: "string" } }, required: ["url"] } },
          footnotes: { type: "array", items: { type: "object", properties: { text: { type: "string" } }, required: ["text"] } },
          endnotes: { type: "array", items: { type: "object", properties: { text: { type: "string" } }, required: ["text"] } },
          bullets: { type: "array", items: { type: ["string", "object"] }, description: "Top-level bulleted list items. Object items may include text and level." },
          numberedList: { type: "array", items: { type: ["string", "object"] }, description: "Top-level numbered list items. Object items may include text and level." },
          sections: {
            type: "array",
            items: {
              type: "object",
              properties: {
                heading: { type: "string" },
                level: { type: "number" },
                body: { type: "string" },
                page: { type: "object", properties: { orientation: { type: "string", enum: ["portrait", "landscape"] }, margins: { type: "object", properties: { top: { type: "number" }, right: { type: "number" }, bottom: { type: "number" }, left: { type: "number" } } }, columns: { type: "object", properties: { count: { type: "number" }, space: { type: "number" }, separator: { type: "boolean" } } } }, description: "Section-specific page setup for generated DOCX output." },
                columns: { type: "object", properties: { count: { type: "number" }, space: { type: "number" }, separator: { type: "boolean" } }, description: "Section-specific Word columns. Use for newspaper/report sections instead of manual layout instructions." },
                footnote: { type: "string", description: "Optional footnote text attached after this section body." },
                endnote: { type: "string", description: "Optional endnote text attached after this section body." },
                revisions: { type: "array", items: { type: "object", properties: { type: { type: "string", enum: ["insert", "delete"] }, text: { type: "string" }, author: { type: "string" }, date: { type: "string" } }, required: ["type", "text"] }, description: "Tracked-change style insertions/deletions to place after this section body." },
                crossReferences: { type: "array", items: { type: "object", properties: { bookmark: { type: "string" }, text: { type: "string" }, fallback: { type: "string" } }, required: ["bookmark"] }, description: "Word REF fields pointing to generated caption bookmarks." },
                links: { type: "array", items: { type: "object", properties: { text: { type: "string" }, label: { type: "string" }, url: { type: "string" } }, required: ["url"] } },
                bullets: { type: "array", items: { type: ["string", "object"] }, description: "Bulleted list items. Object items may include text and level." },
                numberedList: { type: "array", items: { type: ["string", "object"] }, description: "Numbered list items. Object items may include text and level." },
                table: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
                tableCaption: { type: "object", properties: { label: { type: "string" }, text: { type: "string" }, bookmark: { type: "string" } }, description: "Caption to place after this section table using a Word SEQ field and bookmark." },
                images: { type: "array", items: { type: "object", properties: { name: { type: "string" }, type: { type: "string" }, assetId: { type: "string" }, assetName: { type: "string" }, imageUrl: { type: "string" }, base64: { type: "string" }, width: { type: "number" }, height: { type: "number" }, altText: { type: "string" }, caption: { type: "object", properties: { label: { type: "string" }, text: { type: "string" }, bookmark: { type: "string" } } } } } },
              },
            },
          },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "excel_generate_workbook_file",
      description: "Generate a downloadable .xlsx workbook server-side when a standalone workbook artifact is better than live Excel mutation. Supports sheets, typed values, formulas beginning with =, structured table parts, dropdown validations, conditional formats, basic styles, real self-contained chart parts, and real embedded worksheet images from uploaded assets, public image URLs, or base64 data.",
      parameters: {
        type: "object",
        properties: {
          fileName: { type: "string" },
          title: { type: "string", description: "Workbook document title used for generated file metadata." },
          author: { type: "string", description: "Workbook creator/last modified by metadata." },
          subject: { type: "string", description: "Workbook subject metadata." },
          properties: { type: "object", properties: {
            title: { type: "string" }, subject: { type: "string" }, creator: { type: "string" }, author: { type: "string" },
            keywords: { type: ["string", "array"], items: { type: "string" } }, description: { type: "string" }, comments: { type: "string" },
            category: { type: "string" }, company: { type: "string" }, manager: { type: "string" },
          }, description: "Professional workbook document properties for M365 search, governance, and handoff: title, subject, author/creator, keywords, category, company, manager, and description." },
          namedRanges: { type: "array", items: { type: "object", properties: { name: { type: "string" }, sheetName: { type: "string" }, sheetIndex: { type: "number" }, address: { type: "string" }, range: { type: "string" }, ref: { type: "string" }, reference: { type: "string" } }, required: ["name"] }, description: "Workbook-scoped generated named ranges for semantic formulas and downstream automation. Use address/range/ref/reference for the target, such as 'Summary'!$A$1:$B$10." },
          sheets: {
            type: "array",
            items: {
              type: "object",
              properties: {
                name: { type: "string" },
                rows: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
                rowHeights: { type: "array", items: { type: ["number", "object"], properties: { row: { type: "number" }, index: { type: "number" }, min: { type: "number" }, max: { type: "number" }, start: { type: "number" }, end: { type: "number" }, height: { type: "number" }, hidden: { type: "boolean" } } }, description: "Generated worksheet row heights/visibility. Use numbers for sequential row heights or objects for specific rows/ranges." },
                rowsMeta: { type: "array", items: { type: ["number", "object"] }, description: "Alias for generated worksheet row heights/visibility metadata." },
                columns: { type: "array", items: { type: ["number", "object"], properties: { column: { type: ["string", "number"] }, index: { type: "number" }, min: { type: ["string", "number"] }, max: { type: ["string", "number"] }, start: { type: ["string", "number"] }, end: { type: ["string", "number"] }, width: { type: "number" }, hidden: { type: "boolean" } } }, description: "Generated worksheet column widths/visibility. Use numbers for sequential widths or objects for specific columns/ranges." },
                columnWidths: { type: "array", items: { type: ["number", "object"] }, description: "Alias for generated worksheet column widths." },
                merges: { type: "array", items: { type: ["string", "object"], properties: { ref: { type: "string" }, range: { type: "string" } } }, description: "Generated merged cell ranges such as A1:D1 for report headers and section bands." },
                mergeCells: { type: "array", items: { type: ["string", "object"] }, description: "Alias for generated merged cell ranges." },
                namedRanges: { type: "array", items: { type: "object", properties: { name: { type: "string" }, address: { type: "string" }, range: { type: "string" }, ref: { type: "string" }, reference: { type: "string" }, scope: { type: "string", enum: ["sheet", "workbook"] } }, required: ["name"] }, description: "Sheet-scoped generated named ranges. Set scope='workbook' to create a workbook-scoped name from this sheet." },
                freeze: { type: "object", properties: { rows: { type: "number" }, columns: { type: "number" }, cell: { type: "string", description: "Freeze panes at a split cell such as B2." } }, description: "Generated worksheet frozen panes for header rows/columns." },
                freezePanes: { type: "object", properties: { rows: { type: "number" }, columns: { type: "number" }, cell: { type: "string" } }, description: "Alias for freeze." },
                freezeRows: { type: "number", description: "Number of generated worksheet top rows to freeze." },
                freezeColumns: { type: "number", description: "Number of generated worksheet left columns to freeze." },
                zoomScale: { type: "number", description: "Worksheet view zoom percentage, 10-400." },
                orientation: { type: "string", enum: ["portrait", "landscape"], description: "Generated worksheet print orientation." },
                paperSize: { type: "string", enum: ["letter", "legal", "a4"] },
                margins: { type: "object", properties: { top: { type: "number" }, bottom: { type: "number" }, left: { type: "number" }, right: { type: "number" }, header: { type: "number" }, footer: { type: "number" } }, description: "Generated worksheet print margins in inches." },
                topMargin: { type: "number" }, bottomMargin: { type: "number" }, leftMargin: { type: "number" }, rightMargin: { type: "number" },
                printArea: { type: "string", description: "Generated worksheet print area such as A1:G45." },
                autoFilter: { type: ["string", "object"], properties: { ref: { type: "string" }, range: { type: "string" } }, description: "Generated worksheet plain-range AutoFilter ref such as A1:D100. Structured tables already include their own autofilter." },
                filter: { type: ["string", "object"], properties: { ref: { type: "string" }, range: { type: "string" } }, description: "Alias for generated worksheet plain-range AutoFilter." },
                repeatRows: { type: "string", description: "Rows repeated at top when printing, such as 1:2." },
                repeatColumns: { type: "string", description: "Columns repeated at left when printing, such as A:B." },
                fitToPagesWide: { type: "number" }, fitToPagesTall: { type: "number" }, scale: { type: "number" },
                showGridlines: { type: "boolean" }, showHeadings: { type: "boolean" }, centerHorizontally: { type: "boolean" }, centerVertically: { type: "boolean" }, blackAndWhite: { type: "boolean" }, draftMode: { type: "boolean" },
                protect: { type: "boolean", description: "Protect the generated worksheet from accidental edits. This is sheet protection, not encryption." },
                password: { type: "string", description: "Optional Excel sheet-protection password. This is not strong encryption." },
                protection: { type: "object", properties: { protect: { type: "boolean" }, protected: { type: "boolean" }, password: { type: "string" }, allowFormatCells: { type: "boolean" }, allowSort: { type: "boolean" }, allowAutoFilter: { type: "boolean" } }, description: "Generated worksheet protection options for handoff workbooks. Use to prevent accidental edits while optionally allowing formatting, sorting, or filtering." },
                allowFormatCells: { type: "boolean" }, allowSort: { type: "boolean" }, allowAutoFilter: { type: "boolean" },
                comments: { type: "array", items: { type: "object", properties: { address: { type: "string" }, cell: { type: "string" }, text: { type: "string" }, author: { type: "string" }, visible: { type: "boolean" } }, required: ["text"] }, description: "Generated Excel cell comments/notes for review, assumptions, audit notes, and handoff context." },
                notes: { type: "array", items: { type: "object", properties: { address: { type: "string" }, cell: { type: "string" }, text: { type: "string" }, author: { type: "string" }, visible: { type: "boolean" } }, required: ["text"] }, description: "Alias for generated Excel cell comments/notes." },
                links: { type: "array", items: { type: "object", properties: { address: { type: "string" }, cell: { type: "string" }, url: { type: "string" }, target: { type: "string" }, location: { type: "string" }, text: { type: "string" }, label: { type: "string" }, display: { type: "string" }, tooltip: { type: "string" }, screenTip: { type: "string" } }, required: ["address"] }, description: "Generated clickable Excel hyperlinks for source URLs, drill-through links, file links, and handoff references." },
                hyperlinks: { type: "array", items: { type: "object", properties: { address: { type: "string" }, cell: { type: "string" }, url: { type: "string" }, target: { type: "string" }, location: { type: "string" }, text: { type: "string" }, label: { type: "string" }, display: { type: "string" }, tooltip: { type: "string" }, screenTip: { type: "string" } }, required: ["address"] }, description: "Alias for generated clickable Excel hyperlinks." },
                tables: { type: "array", items: { type: "object", properties: { name: { type: "string" }, ref: { type: "string" }, columns: { type: "array", items: { type: "string" } }, style: { type: "string" } } } },
                validations: { type: "array", items: { type: "object", properties: { address: { type: "string" }, values: { type: "array", items: { type: "string" } } } } },
                conditionalFormats: { type: "array", items: { type: "object", properties: { address: { type: "string" }, operator: { type: "string" }, value: { type: ["string", "number"] }, fillColor: { type: "string" } } } },
                charts: { type: "array", items: { type: "object", properties: {
                  title: { type: "string" },
                  chartType: { type: "string", enum: ["bar", "line", "pie", "area", "doughnut", "scatter", "combo"], description: "Generated self-contained chart type. Defaults to bar. Use scatter for numeric X/Y point charts and combo for clustered columns plus line series." },
                  seriesName: { type: "string" },
                  categories: { type: "array", items: { type: "string" }, description: "Preferred category labels for multi-series charts, such as quarters, months, products, or scenarios." },
                  series: { type: "array", items: { type: "object", properties: { name: { type: "string" }, chartType: { type: "string", enum: ["bar", "column", "line"], description: "For combo charts, choose whether this series renders as columns or line." }, values: { type: "array", items: { type: "number" } }, points: { type: "array", items: { type: "array", items: { type: "number" } }, description: "For scatter charts, numeric [x, y] points." } }, required: ["name"] }, description: "Preferred multi-series chart data. Use values for category/combo charts or points for scatter charts." },
                  lineSeries: { type: "array", items: { type: "string" }, description: "For combo charts, series names to render as line series." },
                  lineSeriesStartIndex: { type: "number", description: "For combo charts, zero-based series index where line rendering begins. Defaults to the final series." },
                  points: { type: "array", items: { type: "array", items: { type: "number" } }, description: "For scatter charts, single-series numeric [x, y] points." },
                  values: { type: "array", items: { type: "array", items: { type: ["string", "number"] } }, description: "Legacy single-series rows of [category, numeric value], or numeric [x, y] points when chartType is scatter. Use categories + series for comparisons." },
                  categoryAxisTitle: { type: "string", description: "Optional category/X axis title for bar, line, and area charts." },
                  valueAxisTitle: { type: "string", description: "Optional value/Y axis title for bar, line, and area charts." },
                  valueFormat: { type: "string", description: "Optional chart value number format, such as $#,##0 or 0%." },
                  xValueFormat: { type: "string", description: "Optional scatter X-axis number format." },
                  yValueFormat: { type: "string", description: "Optional scatter Y-axis number format." },
                  dataLabels: { type: "boolean", description: "Show value data labels on the chart." },
                  scatterStyle: { type: "string", enum: ["marker", "line", "lineMarker", "smooth", "smoothMarker"], description: "Scatter chart style. Defaults to marker." },
                  colors: { type: "array", items: { type: "string" }, description: "Optional hex colors for chart series/fills." },
                  legendPosition: { type: "string", enum: ["right", "left", "top", "bottom", "none"] },
                  barDirection: { type: "string", enum: ["vertical", "horizontal"], description: "For bar charts, choose vertical columns or horizontal bars." },
                  showGridLines: { type: "boolean" },
                  holeSize: { type: "number", description: "For doughnut charts, hole size from 10 to 90. Defaults to 50." },
                  cell: { type: "string" }, startCell: { type: "string" }, width: { type: "number" }, height: { type: "number" },
                } } },
                images: { type: "array", items: { type: "object", properties: { name: { type: "string" }, type: { type: "string" }, assetId: { type: "string" }, assetName: { type: "string" }, imageUrl: { type: "string" }, base64: { type: "string" }, cell: { type: "string" }, startCell: { type: "string" }, width: { type: "number" }, height: { type: "number" }, altText: { type: "string" } } } },
              },
              required: ["name", "rows"],
            },
          },
          rows: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "powerpoint_generate_deck_file",
      description: "Generate a downloadable .pptx file server-side when live PowerPoint APIs cannot reliably create the needed deck features. Supports slide title/body/subtitle, speaker notes, background colors, real shapes, real table graphic frames, clickable hyperlink text, editable chart parts, real embedded images, and template-preserving generation from an uploaded/source PPTX template asset with title/body/subtitle plus image/chart/table placeholder placement when the source layout exposes placeholders.",
      parameters: {
        type: "object",
        properties: {
          fileName: { type: "string", description: "Output file name ending in .pptx." },
          title: { type: "string", description: "Deck title." },
          footer: { type: "string", description: "Optional default footer text to place at the bottom of generated slides." },
          footerText: { type: "string", description: "Alias for footer." },
          dateText: { type: "string", description: "Optional default date/version text to place at the bottom of generated slides." },
          date: { type: "string", description: "Alias for dateText." },
          showDate: { type: "boolean", description: "Whether to show dateText/date when supplied. Defaults to true when text exists." },
          showFooter: { type: "boolean", description: "Whether to show footer/footerText when supplied. Defaults to true when text exists." },
          showSlideNumber: { type: "boolean", description: "Show generated slide numbers as editable text boxes." },
          showSlideNumbers: { type: "boolean", description: "Alias for showSlideNumber." },
          slideNumberFormat: { type: "string", description: "Slide number format using {n} and {total}, such as {n}/{total} or Slide {n}. Defaults to {n}." },
          slideNumberText: { type: "string", description: "Override text/template for generated slide number labels using {n} and {total}." },
          confidentialityLabel: { type: "string", description: "Optional default confidentiality/classification label, such as Confidential or Internal Use Only." },
          footerColor: { type: "string", description: "Default footer/date/slide-number text color." },
          footerFontSize: { type: "number", description: "Default footer/date/slide-number font size in DrawingML hundredths of a point, e.g. 850." },
          templateAssetId: { type: "string", description: "Optional uploaded PowerPoint template asset id to preserve masters, layouts, theme, relationships, fonts, and reusable media while generating new slides." },
          templateAssetName: { type: "string", description: "Optional uploaded PowerPoint template filename to preserve as the source template." },
          templateBase64: { type: "string", description: "Optional base64 or data URL of a .pptx template package. Prefer templateAssetId/templateAssetName when the file was attached in the pane." },
          slides: {
            type: "array",
            items: {
              type: "object",
              properties: {
                title: { type: "string" },
                subtitle: { type: "string", description: "Optional subtitle/kicker text. Template-preserving generation fills subtitle placeholders when present." },
                layoutName: { type: "string", description: "Optional source-template layout name to use for this slide, such as Title Slide, Section Header, Chart Slide, or Two Content. Use names from attached PPTX context when available." },
                layoutType: { type: "string", description: "Optional source-template layout type to use, such as title, section, chart, table, picture, or content." },
                layoutIndex: { type: "number", description: "Optional zero-based source-template layout index. Prefer layoutName when the context exposes names." },
                body: { type: "string" },
                notes: { type: "string" },
                backgroundColor: { type: "string", description: "Optional slide background hex color such as #E0F2FE." },
                footer: { type: "string", description: "Optional per-slide footer override." },
                footerText: { type: "string", description: "Alias for per-slide footer." },
                dateText: { type: "string", description: "Optional per-slide date/version text override." },
                date: { type: "string", description: "Alias for dateText." },
                showDate: { type: "boolean" },
                showFooter: { type: "boolean" },
                showSlideNumber: { type: "boolean" },
                showSlideNumbers: { type: "boolean" },
                slideNumberFormat: { type: "string", description: "Per-slide slide-number format using {n} and {total}." },
                slideNumberText: { type: "string", description: "Per-slide slide-number text/template using {n} and {total}." },
                confidentialityLabel: { type: "string", description: "Optional per-slide confidentiality/classification label." },
                footerColor: { type: "string" },
                footerFontSize: { type: "number" },
                shapes: {
                  type: "array",
                  items: {
                    type: "object",
                    properties: {
                      shapeType: { type: "string", enum: ["rectangle", "roundedRectangle", "oval", "triangle", "diamond", "pentagon", "hexagon", "cloud"] },
                      text: { type: "string" },
                      left: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      top: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      width: { type: "number", description: "Optional explicit width in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      height: { type: "number", description: "Optional explicit height in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      fillColor: { type: "string" },
                      lineColor: { type: "string" },
                      fontSize: { type: "number" },
                    },
                  },
                },
                tables: {
                  type: "array",
                  items: {
                    type: "object",
                    properties: {
                      name: { type: "string" },
                      values: { type: "array", items: { type: "array", items: { type: ["string", "number", "boolean", "null", "object"], properties: { text: { type: ["string", "number", "boolean", "null"] }, value: { type: ["string", "number", "boolean", "null"] }, fillColor: { type: "string" }, textColor: { type: "string" }, bold: { type: "boolean" }, fontSize: { type: "number" }, align: { type: "string", enum: ["left", "center", "right", "justify"] }, verticalAlign: { type: "string", enum: ["top", "middle", "bottom"] } } } } },
                      left: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      top: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      width: { type: "number", description: "Optional explicit width in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      height: { type: "number", description: "Optional explicit height in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      headerFillColor: { type: "string" },
                      headerTextColor: { type: "string" },
                      bodyFillColor: { type: "string" },
                      bandFillColor: { type: "string", description: "Alternating row fill color for generated PowerPoint tables." },
                      alternateRowFillColor: { type: "string", description: "Alias for bandFillColor." },
                      textColor: { type: "string" },
                      borderColor: { type: "string" },
                      borderWidth: { type: "number" },
                      fontSize: { type: "number" },
                      headerFontSize: { type: "number" },
                      headerBold: { type: "boolean" },
                      bold: { type: "boolean" },
                      align: { type: "string", enum: ["left", "center", "right", "justify"] },
                      headerAlign: { type: "string", enum: ["left", "center", "right", "justify"] },
                      verticalAlign: { type: "string", enum: ["top", "middle", "bottom"] },
                      firstRow: { type: "boolean", description: "Whether the first row should be styled as a header. Defaults to true." },
                      bandRows: { type: "boolean", description: "Whether alternating body rows should be banded. Defaults to true." },
                      columnWidths: { type: "array", items: { type: "number" }, description: "Optional column widths in points." },
                      rowHeights: { type: "array", items: { type: "number" }, description: "Optional row heights in points." },
                    },
                  },
                },
                links: { type: "array", items: { type: "object", properties: { text: { type: "string" }, label: { type: "string" }, url: { type: "string" }, left: { type: "number" }, top: { type: "number" }, width: { type: "number" }, height: { type: "number" }, fontSize: { type: "number" } }, required: ["url"] } },
                charts: {
                  type: "array",
                  items: {
                    type: "object",
                    properties: {
                      title: { type: "string" },
                      chartType: { type: "string", enum: ["bar", "line", "pie", "area", "doughnut", "scatter", "combo"], description: "Generated editable chart type. Defaults to bar. Use scatter for numeric X/Y point charts and combo for clustered columns plus line series." },
                      seriesName: { type: "string", description: "Legacy single-series name when using values." },
                      categories: { type: "array", items: { type: "string" }, description: "Preferred category labels for multi-series charts, such as quarters, months, products, or scenarios." },
                      series: { type: "array", items: { type: "object", properties: { name: { type: "string" }, chartType: { type: "string", enum: ["bar", "column", "line"], description: "For combo charts, choose whether this series renders as columns or line." }, values: { type: "array", items: { type: "number" } }, points: { type: "array", items: { type: "array", items: { type: "number" } }, description: "For scatter charts, numeric [x, y] points." } }, required: ["name"] }, description: "Preferred multi-series chart data. Use values for category/combo charts or points for scatter charts." },
                      lineSeries: { type: "array", items: { type: "string" }, description: "For combo charts, series names to render as line series." },
                      lineSeriesStartIndex: { type: "number", description: "For combo charts, zero-based series index where line rendering begins. Defaults to the final series." },
                      points: { type: "array", items: { type: "array", items: { type: "number" } }, description: "For scatter charts, single-series numeric [x, y] points." },
                      values: { type: "array", items: { type: "array", items: { type: ["string", "number"] } }, description: "Legacy single-series rows of [category, numeric value], or numeric [x, y] points when chartType is scatter. Use categories + series for comparisons." },
                      categoryAxisTitle: { type: "string", description: "Optional category/X axis title for bar, line, and area charts." },
                      valueAxisTitle: { type: "string", description: "Optional value/Y axis title for bar, line, and area charts." },
                      valueFormat: { type: "string", description: "Optional chart value number format, such as $#,##0 or 0%." },
                      xValueFormat: { type: "string", description: "Optional scatter X-axis number format." },
                      yValueFormat: { type: "string", description: "Optional scatter Y-axis number format." },
                      dataLabels: { type: "boolean", description: "Show value data labels on the chart." },
                      scatterStyle: { type: "string", enum: ["marker", "line", "lineMarker", "smooth", "smoothMarker"], description: "Scatter chart style. Defaults to marker." },
                      colors: { type: "array", items: { type: "string" }, description: "Optional hex colors for chart series/fills." },
                      legendPosition: { type: "string", enum: ["right", "left", "top", "bottom"], description: "Optional legend position. Defaults to right." },
                      barDirection: { type: "string", enum: ["vertical", "horizontal"], description: "For bar charts, choose vertical columns or horizontal bars." },
                      showGridLines: { type: "boolean", description: "Show or hide major value gridlines. Defaults to true." },
                      holeSize: { type: "number", description: "For doughnut charts, hole size from 10 to 90. Defaults to 50." },
                      left: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      top: { type: "number", description: "Optional explicit position in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      width: { type: "number", description: "Optional explicit width in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                      height: { type: "number", description: "Optional explicit height in points. Omit when using a source PPTX template and you want image/table/chart placeholders to determine placement." },
                    },
                    required: [],
                  },
                },
                images: {
                  type: "array",
                  items: {
                    type: "object",
                    properties: {
                      name: { type: "string" },
                      type: { type: "string" },
                      assetId: { type: "string", description: "Uploaded image asset id." },
                      assetName: { type: "string", description: "Uploaded image filename." },
                      imageUrl: { type: "string", description: "Public http(s) image URL to fetch and embed." },
                      base64: { type: "string", description: "Optional base64 image bytes or data URL." },
                      left: { type: "number" },
                      top: { type: "number" },
                      width: { type: "number" },
                      height: { type: "number" },
                      pixelWidth: { type: "number", description: "Original image pixel width when known; used for fit/fill crop math." },
                      pixelHeight: { type: "number", description: "Original image pixel height when known; used for fit/fill crop math." },
                      fit: { type: "string", enum: ["stretch", "fit", "contain", "fill", "cover"], description: "Generated PowerPoint image sizing mode. Use fit/contain to preserve the whole image inside the box, fill/cover to center-crop into the box, and stretch only when distortion is acceptable." },
                      sizing: { type: "string", enum: ["stretch", "fit", "contain", "fill", "cover"], description: "Alias for fit." },
                      crop: { type: "object", properties: { left: { type: "number" }, right: { type: "number" }, top: { type: "number" }, bottom: { type: "number" } }, description: "Explicit crop percentages. Values can be 0-1 or 0-100." },
                      cropLeft: { type: "number" },
                      cropRight: { type: "number" },
                      cropTop: { type: "number" },
                      cropBottom: { type: "number" },
                      altText: { type: "string" },
                    },

                  },
                },
              },
              required: ["title"],
            },
          },
        },
        required: ["slides"],
      },
    },
  },
] as const;

function result(call: ToolCallRequest, ok: boolean, content: string): ToolCallResult {
  return { id: call.id, name: call.name, ok, content };
}

function demoTheme(value: unknown) {
  const theme = typeof value === "string" && value.trim() ? value.trim() : "CTRL Office Power User Showcase";
  return theme.slice(0, 120);
}

function demoAudience(value: unknown) {
  const audience = typeof value === "string" && value.trim() ? value.trim() : "executive";
  return audience.slice(0, 40);
}

function excelDemoPayload(theme: string, audience: string) {
  const rows = [
    ["Month", "Revenue", "Pipeline", "Risk", "Status"],
    ["Jan", 1.2, 2.4, 0.18, "On track"],
    ["Feb", 1.4, 2.8, 0.16, "On track"],
    ["Mar", 1.7, 3.1, 0.14, "Strong"],
    ["Apr", 1.5, 3.4, 0.21, "Watch"],
    ["May", 1.9, 3.8, 0.17, "Strong"],
    ["Jun", 2.2, 4.1, 0.13, "Strong"],
  ];
  return {
    fileName: "CTRL Excel Feature Demo.xlsx",
    title: `${theme} - Excel Demo`,
    subject: `Prepared ${audience} demonstration workbook`,
    author: "CTRL Add-in Demo Agent",
    properties: { company: "CTRL", category: "Demo", keywords: "Excel, AI, Office add-in, charts, filters, comments" },
    namedRanges: [{ name: "Demo_Revenue_Table", sheetName: "Showcase", address: "A4:E10" }],
    sheets: [{
      name: "Showcase",
      rows: [[`${theme}: Excel power-user demo`], [`Audience: ${audience}`], [], ...rows, [], ["Demo coverage", "Tables, formulas, charts, comments, hyperlinks, filters, frozen panes, print layout, protection"]],
      merges: ["A1:E1", "A2:E2"],
      columns: [{ min: 1, max: 1, width: 16 }, { min: 2, max: 4, width: 14 }, { min: 5, max: 5, width: 18 }],
      rowHeights: [28, 20, 8, 24],
      freeze: { rows: 4 },
      autoFilter: "A4:E10",
      tables: [{ name: "DemoMetrics", ref: "A4:E10", columns: ["Month", "Revenue", "Pipeline", "Risk", "Status"], style: "TableStyleMedium2" }],
      validations: [{ address: "E5:E10", values: ["On track", "Strong", "Watch"] }],
      conditionalFormats: [{ address: "D5:D10", operator: "greaterThan", value: 0.2, fillColor: "#FEE2E2" }],
      comments: [{ address: "D8", text: "The demo agent flags risk with conditional formatting and comments for review handoff.", author: "CTRL Demo" }],
      hyperlinks: [{ address: "A13", text: "CTRL generated workbook demo", url: "https://example.test/ctrl-demo", tooltip: "Example demo link" }],
      charts: [{ title: "Revenue and pipeline", chartType: "combo", categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"], series: [{ name: "Revenue", values: [1.2, 1.4, 1.7, 1.5, 1.9, 2.2], chartType: "bar" }, { name: "Pipeline", values: [2.4, 2.8, 3.1, 3.4, 3.8, 4.1], chartType: "line" }], valueFormat: "$0.0M", categoryAxisTitle: "Month", valueAxisTitle: "USD", dataLabels: true, legendPosition: "bottom", cell: "G4", width: 520, height: 280 }],
      printArea: "A1:M18",
      orientation: "landscape",
      fitToPagesWide: 1,
      fitToPagesTall: 1,
      showGridlines: false,
      protect: true,
      allowAutoFilter: true,
      allowSort: true,
    }],
  };
}

function wordDemoPayload(theme: string, audience: string) {
  return {
    fileName: "CTRL Word Feature Demo.docx",
    title: `${theme} - Word Demo`,
    author: "CTRL Add-in Demo Agent",
    subject: `Prepared ${audience} demonstration document`,
    header: "CTRL Add-in Demonstration",
    footer: "Generated with CTRL BYOK Office Add-in",
    tableOfContents: true,
    orientation: "portrait",
    margins: { top: 0.7, bottom: 0.7, left: 0.8, right: 0.8 },
    paragraphs: ["This prepared demo shows how CTRL can draft, structure, annotate, and package a professional Word deliverable without leaving Office."],
    numberedList: ["Read current Office context", "Use BYOK model/tool calling", "Generate a polished Word artifact"],
    sections: [
      { heading: "Executive summary", body: `For a ${audience} audience, CTRL can turn chat instructions into document structure: headings, sections, tables, links, comments, captions, footnotes, endnotes, and revision markup.`, bullets: ["Real headings and lists", "Report headers and footers", "Review comments and tracked-change style markup"], footnote: "Demo footnote showing source/context support." },
      { heading: "Capability matrix", body: "The table below is generated as a real Word table, not a screenshot.", table: [["Area", "Demo feature"], ["Authoring", "Headings, sections, lists, tables"], ["Review", "Comments and redline-style markup"], ["Research", "Footnotes, links, M365/file context"]], tableCaption: { label: "Table", text: "CTRL Word feature coverage", bookmark: "tbl_ctrl_word_demo" }, comments: [{ text: "This comment demonstrates review handoff context.", author: "CTRL Demo" }] },
      { heading: "Next-level workflow", body: "The model can use Office tools first, then switch to generated DOCX when the live runtime lacks a ribbon-level feature.", links: [{ text: "Example reference", url: "https://example.test/ctrl-word-demo" }], revisions: [{ type: "insert", text: "Inserted recommendation: use generated artifacts for repeatable board-pack output.", author: "CTRL Demo" }, { type: "delete", text: "Manual formatting instructions only.", author: "CTRL Demo" }], endnote: "Demo endnote for appendix-style context." },
    ],
  };
}

function powerpointDemoPayload(theme: string, audience: string) {
  return {
    fileName: "CTRL PowerPoint Feature Demo.pptx",
    title: `${theme} - PowerPoint Demo`,
    footer: "CTRL Add-in Feature Demo",
    dateText: "Prepared demo",
    confidentialityLabel: audience === "sales" ? "Customer-ready" : "Internal demo",
    showSlideNumber: true,
    slideNumberFormat: "{n}/{total}",
    footerColor: "64748B",
    slides: [
      { title: "CTRL inside PowerPoint", subtitle: "Prepared feature showcase", body: "Ask for a demo and CTRL can generate a polished deck with real objects, charts, tables, notes, and branded layout polish.", backgroundColor: "F8FAFC", notes: "Open by explaining that this is generated as editable PowerPoint XML, not screenshots.", shapes: [{ shapeType: "roundedRectangle", text: "BYOK + tools + Office context", left: 70, top: 310, width: 580, height: 62, fillColor: "DBEAFE", lineColor: "2563EB", fontSize: 1500 }] },
      { title: "Real slide objects", body: "The demo uses editable shapes, styled tables, hyperlinks, and chart parts so users can keep working in PowerPoint.", tables: [{ name: "Feature Matrix", values: [["Feature", "Shown here"], ["Styled tables", { text: "Yes", fillColor: "DCFCE7", textColor: "166534", bold: true }], ["Editable charts", { text: "Yes", fillColor: "DCFCE7", textColor: "166534", bold: true }], ["Footer/date/slide numbers", { text: "Yes", fillColor: "DCFCE7", textColor: "166534", bold: true }]], left: 70, top: 165, width: 590, height: 190, headerFillColor: "1F4E79", headerTextColor: "FFFFFF", bandFillColor: "EFF6FF", borderColor: "94A3B8", columnWidths: [340, 250] }], links: [{ text: "Demo link", url: "https://example.test/ctrl-powerpoint-demo", left: 70, top: 385, width: 180, height: 28 }] },
      { title: "Editable charts", body: "Generated decks can package chart parts with workbook-backed chart data for professional storytelling.", charts: [{ title: "Demo value by workflow", chartType: "combo", categories: ["Draft", "Analyze", "Review", "Present"], series: [{ name: "Time saved", values: [30, 45, 25, 40], chartType: "bar" }, { name: "Quality lift", values: [15, 28, 35, 42], chartType: "line" }], valueFormat: "0", dataLabels: true, categoryAxisTitle: "Workflow", valueAxisTitle: "Index", legendPosition: "bottom", colors: ["2563EB", "7C3AED"], left: 90, top: 170, width: 540, height: 260 }], notes: "Call out that the chart is an editable chart object, not an image." },
      { title: "Fallbacks that still produce real files", body: "If a live Office API cannot perform a ribbon-level action, CTRL should switch to generated Office artifacts instead of returning manual instructions.", shapes: [{ shapeType: "diamond", text: "Live tool?", left: 110, top: 185, width: 120, height: 90, fillColor: "E0F2FE", lineColor: "0284C7" }, { shapeType: "roundedRectangle", text: "Generated file fallback", left: 330, top: 190, width: 230, height: 80, fillColor: "F3E8FF", lineColor: "7C3AED" }], notes: "This slide demonstrates the product philosophy: no placeholders when a real generated file can satisfy the ask." },
    ],
  };
}

async function postGeneratedDemo(kind: "pptx" | "docx" | "xlsx", payload: Record<string, unknown>) {
  const path = kind === "pptx" ? "/api/generated/pptx" : kind === "docx" ? "/api/generated/docx" : "/api/generated/xlsx";
  const response = await fetch(path, { method: "POST", headers: { "content-type": "application/json", accept: "application/json" }, body: JSON.stringify(payload) });
  const text = await response.text();
  return { kind, ok: response.ok, response: withArtifactLink(kind, text) };
}

function tryJson(text: string) {
  try { return JSON.parse(text || "null"); } catch { return text; }
}

function artifactLabel(kind: "pptx" | "docx" | "xlsx") {
  if (kind === "pptx") return "PowerPoint artifact";
  if (kind === "docx") return "Word artifact";
  return "Excel artifact";
}

function withArtifactLink(kind: "pptx" | "docx" | "xlsx", text: string) {
  const parsed = tryJson(text);
  if (!parsed || typeof parsed !== "object" || !("downloadUrl" in parsed)) return text;
  const downloadUrl = String((parsed as any).downloadUrl || "").trim();
  if (!downloadUrl) return text;
  return JSON.stringify({ ...(parsed as any), markdownLink: `[Download the ${artifactLabel(kind)}](${downloadUrl})`, note: "Show markdownLink exactly as provided; do not insert a space between ] and (." });
}

type DemoSurface = "excel" | "word" | "powerpoint";
type DemoMode = "auto" | "live" | "artifact";

function demoMode(value: unknown): DemoMode {
  return value === "live" || value === "artifact" || value === "auto" ? value : "auto";
}

function demoCall(name: string, args: Record<string, unknown>, suffix: string): ToolCallRequest {
  return { id: `demo-${Date.now()}-${suffix}`, name, arguments: args };
}

async function runLiveDemoStep(host: DemoSurface, steps: ToolCallResult[], name: string, args: Record<string, unknown>, suffix: string, required = true) {
  const step = await executeOfficeTool(host, demoCall(name, args, suffix));
  steps.push(step);
  if (required && !step.ok) throw new Error(step.content || `${name} failed.`);
  return step;
}

function compactDemoSteps(steps: ToolCallResult[]) {
  return steps.map((step) => ({ tool: step.name, ok: step.ok, summary: step.content }));
}

async function createLiveExcelDemo(theme: string, audience: string) {
  const sheetName = "CTRL Demo";
  const rows = [
    ["Month", "Revenue", "Pipeline", "Risk", "Status"],
    ["Jan", 1.2, 2.4, 0.18, "On track"],
    ["Feb", 1.4, 2.8, 0.16, "On track"],
    ["Mar", 1.7, 3.1, 0.14, "Strong"],
    ["Apr", 1.5, 3.4, 0.21, "Watch"],
    ["May", 1.9, 3.8, 0.17, "Strong"],
    ["Jun", 2.2, 4.1, 0.13, "Strong"],
  ];
  const steps: ToolCallResult[] = [];
  await runLiveDemoStep("excel", steps, "excel_add_worksheet", { name: sheetName }, "excel-sheet");
  await runLiveDemoStep("excel", steps, "excel_write_range", { sheetName, address: "A1", values: [[`${theme}: Excel live demo`], [`Audience: ${audience}`], [], ...rows, [], ["What this shows", "Live worksheet creation, range writing, tables, formatting, validation, conditional formatting, comments, named ranges, freeze panes, and a native chart."]] }, "excel-data");
  await runLiveDemoStep("excel", steps, "excel_create_table", { sheetName, sourceAddress: "A4:E10", name: "CTRL_Demo_Metrics", hasHeaders: true, style: "TableStyleMedium2" }, "excel-table", false);
  await runLiveDemoStep("excel", steps, "excel_format_range", { sheetName, address: "A1:E1", fillColor: "#1f4e79", fontColor: "#ffffff", bold: true, fontSize: 16, horizontalAlignment: "Center", wrapText: true }, "excel-title", false);
  await runLiveDemoStep("excel", steps, "excel_format_range", { sheetName, address: "A4:E4", fillColor: "#dbeafe", bold: true, wrapText: true }, "excel-header", false);
  await runLiveDemoStep("excel", steps, "excel_add_data_validation", { sheetName, address: "E5:E10", listValues: ["On track", "Strong", "Watch"], promptTitle: "Demo status", promptMessage: "Choose a status value." }, "excel-validation", false);
  await runLiveDemoStep("excel", steps, "excel_apply_conditional_format", { sheetName, address: "D5:D10", ruleType: "greaterThan", value: 0.2, fillColor: "#fee2e2", fontColor: "#991b1b" }, "excel-conditional", false);
  await runLiveDemoStep("excel", steps, "excel_freeze_panes", { sheetName, rows: 4 }, "excel-freeze", false);
  await runLiveDemoStep("excel", steps, "excel_add_comment", { sheetName, address: "D8", text: "CTRL can add review notes while preparing live analysis." }, "excel-comment", false);
  await runLiveDemoStep("excel", steps, "excel_set_named_range", { sheetName, name: "CTRL_Demo_Metrics", address: "A4:E10" }, "excel-name", false);
  await runLiveDemoStep("excel", steps, "excel_create_chart", { sheetName, sourceAddress: "A4:C10", chartType: "lineMarkers", title: "Revenue and pipeline", startCell: "G4", endCell: "M18" }, "excel-chart", false);
  return { ok: steps.every((step) => step.ok), surface: "excel", mode: "live", message: "Created a live CTRL Excel demo on a new worksheet named CTRL Demo.", steps: compactDemoSteps(steps) };
}

async function createLiveWordDemo(theme: string, audience: string) {
  const steps: ToolCallResult[] = [];
  await runLiveDemoStep("word", steps, "word_insert_heading", { text: `${theme}: CTRL Word live demo`, level: 1, location: "end" }, "word-heading");
  await runLiveDemoStep("word", steps, "word_insert_text", { location: "end", text: `\nThis live demo shows CTRL working inside Word for a ${audience} audience: drafting structure, creating tables, adding review notes, and applying report polish without switching apps.\n` }, "word-intro");
  await runLiveDemoStep("word", steps, "word_insert_heading", { text: "Capability matrix", level: 2, location: "end" }, "word-matrix-heading", false);
  await runLiveDemoStep("word", steps, "word_insert_table", { location: "end", style: "Grid Table 4 - Accent 1", values: [["Workflow", "Live capability"], ["Authoring", "Headings and structured prose"], ["Analysis", "Tables grounded in context"], ["Review", "Comments and controlled fields"], ["Packaging", "Headers, footers, and generated DOCX fallback"]] }, "word-table", false);
  await runLiveDemoStep("word", steps, "word_insert_comment", { text: "CTRL demo note: the assistant can leave review context directly in the document when the runtime supports comments." }, "word-comment", false);
  await runLiveDemoStep("word", steps, "word_set_header_footer", { header: "CTRL Add-in Live Demo", footer: "Created inside the active Word document" }, "word-header-footer", false);
  await runLiveDemoStep("word", steps, "word_insert_content_control", { title: "Next step", tag: "ctrl-demo-next-step", placeholderText: "Ask CTRL to turn this into a reusable template.", location: "end" }, "word-control", false);
  return { ok: steps.every((step) => step.ok), surface: "word", mode: "live", message: "Created a live CTRL Word demo by appending structured content to the current document.", steps: compactDemoSteps(steps) };
}

async function createLivePowerPointDemo(theme: string, audience: string) {
  const steps: ToolCallResult[] = [];
  await runLiveDemoStep("powerpoint", steps, "powerpoint_create_slides", { slides: [{ title: "CTRL inside PowerPoint", body: `Live demo for ${audience}: CTRL can create slides, tables, text boxes, shapes, backgrounds, and layout polish directly in the current deck.` }] }, "ppt-slide-1");
  await runLiveDemoStep("powerpoint", steps, "powerpoint_set_slide_background", { slidePosition: "last", color: "#f8fafc" }, "ppt-bg-1", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_add_shape", { slidePosition: "last", shapeType: "roundedRectangle", text: "BYOK + tools + Office context", left: 72, top: 300, width: 560, height: 56, fillColor: "#dbeafe", lineColor: "#2563eb" }, "ppt-pill", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_create_slides", { slides: [{ title: "Native objects, not placeholders", body: "This demo adds editable PowerPoint objects through the task pane runtime, then uses generated PPTX only when a live API is unavailable." }] }, "ppt-slide-2");
  await runLiveDemoStep("powerpoint", steps, "powerpoint_add_table", { slidePosition: "last", left: 64, top: 150, width: 590, height: 180, headerFillColor: "#1f4e79", values: [["Feature", "Live demo"], ["Slides", "Created in current deck"], ["Tables", "Editable slide grid"], ["Shapes", "Styled visual objects"], ["Fallbacks", "Generated PPTX when needed"]] }, "ppt-table", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_create_slides", { slides: [{ title: "Workflow coverage", body: "Research, write, structure, format, and package presentation work without leaving PowerPoint." }] }, "ppt-slide-3");
  await runLiveDemoStep("powerpoint", steps, "powerpoint_add_textbox", { slidePosition: "last", text: "Next: ask CTRL to turn company notes, uploaded decks, or web research into a branded presentation.", left: 76, top: 230, width: 560, height: 80, fontSize: 20, fontColor: "#1f2937" }, "ppt-next", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_add_shape", { slidePosition: "last", shapeType: "roundedRectangle", text: "Ask", left: 120, top: 150, width: 90, height: 70, fillColor: "#e0f2fe", lineColor: "#0284c7" }, "ppt-shape-1", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_add_shape", { slidePosition: "last", shapeType: "roundedRectangle", text: "Create", left: 305, top: 150, width: 120, height: 70, fillColor: "#f3e8ff", lineColor: "#7c3aed" }, "ppt-shape-2", false);
  await runLiveDemoStep("powerpoint", steps, "powerpoint_arrange_shapes", { slidePosition: "last", shapeNumbers: [3, 4], action: "alignMiddle" }, "ppt-arrange", false);
  return { ok: steps.every((step) => step.ok), surface: "powerpoint", mode: "live", message: "Created a live CTRL PowerPoint demo by adding editable slides and objects to the current deck.", note: "Speaker notes and full generated chart parts remain available through artifact mode because the live PowerPoint runtime may not expose those APIs.", steps: compactDemoSteps(steps) };
}

async function createLiveDemo(surface: DemoSurface, theme: string, audience: string) {
  if (surface === "excel") return createLiveExcelDemo(theme, audience);
  if (surface === "word") return createLiveWordDemo(theme, audience);
  return createLivePowerPointDemo(theme, audience);
}

async function createArtifactDemos(surface: "all" | DemoSurface, theme: string, audience: string) {
  const jobs: Array<Promise<{ kind: "pptx" | "docx" | "xlsx"; ok: boolean; response: string }>> = [];
  if (surface === "all" || surface === "excel") jobs.push(postGeneratedDemo("xlsx", excelDemoPayload(theme, audience)));
  if (surface === "all" || surface === "word") jobs.push(postGeneratedDemo("docx", wordDemoPayload(theme, audience)));
  if (surface === "all" || surface === "powerpoint") jobs.push(postGeneratedDemo("pptx", powerpointDemoPayload(theme, audience)));
  const outputs = await Promise.all(jobs);
  const ok = outputs.every((output) => output.ok);
  return { ok, surface, mode: "artifact", theme, audience, generated: outputs.map((output) => ({ kind: output.kind, ok: output.ok, result: tryJson(output.response) })) };
}

async function demoShowcaseTool(host: OfficeHost, call: ToolCallRequest) {
  if (call.name !== "ctrl_create_demo_showcase") return null;
  const requested = typeof call.arguments.surface === "string" ? call.arguments.surface : "current";
  const surface = requested === "current" ? (host === "unknown" ? "all" : host) : requested;
  const mode = demoMode(call.arguments.mode);
  const theme = demoTheme(call.arguments.theme);
  const audience = demoAudience(call.arguments.audience);
  if (surface !== "all" && surface !== "excel" && surface !== "word" && surface !== "powerpoint") return result(call, false, "Unknown demo surface. Use current, excel, word, powerpoint, or all.");

  if (mode === "artifact" || surface === "all" || host === "unknown") {
    const artifacts = await createArtifactDemos(surface as "all" | DemoSurface, theme, audience);
    return result(call, artifacts.ok, JSON.stringify(artifacts));
  }

  if (host !== surface) {
    if (mode === "live") return result(call, false, `Live ${surface} demos can only run while the add-in is open in ${surface}. Use mode=artifact for a downloadable demo file, or surface=current in the active Office app.`);
    const artifacts = await createArtifactDemos(surface as DemoSurface, theme, audience);
    return result(call, artifacts.ok, JSON.stringify({ ...artifacts, fallbackReason: `The active Office host is ${host}, so CTRL cannot live-edit ${surface} from here.` }));
  }

  try {
    const live = await createLiveDemo(surface as DemoSurface, theme, audience);
    return result(call, true, JSON.stringify({ ...live, theme, audience }));
  } catch (error: any) {
    if (mode === "live") return result(call, false, `Could not create the live ${surface} demo: ${error?.message || String(error)}`);
    const artifacts = await createArtifactDemos(surface as DemoSurface, theme, audience);
    return result(call, artifacts.ok, JSON.stringify({ ...artifacts, fallbackReason: `The live ${surface} demo could not run in this Office runtime: ${error?.message || String(error)}` }));
  }
}

function parseArgs(value: unknown) {
  if (!value) return {};
  if (typeof value === "string") {
    try { return JSON.parse(value); } catch { return { raw: value }; }
  }
  if (typeof value === "object") return value as Record<string, unknown>;
  return { value };
}

async function webTool(call: ToolCallRequest) {
  if (call.name === "web_search") {
    const query = typeof call.arguments.query === "string" ? call.arguments.query.trim() : "";
    if (!query) return result(call, false, "web_search requires query.");
    const response = await fetch(`/api/web-search?q=${encodeURIComponent(query)}`);
    return result(call, response.ok, await response.text());
  }

  if (call.name === "web_fetch") {
    const url = typeof call.arguments.url === "string" ? call.arguments.url.trim() : "";
    if (!url) return result(call, false, "web_fetch requires url.");
    const response = await fetch(`/api/web-fetch?url=${encodeURIComponent(url)}`);
    return result(call, response.ok, await response.text());
  }

  if (call.name === "web_image_search") {
    const query = typeof call.arguments.query === "string" ? call.arguments.query.trim() : "";
    const count = typeof call.arguments.count === "number" ? call.arguments.count : 8;
    if (!query) return result(call, false, "web_image_search requires query.");
    const response = await fetch(`/api/web-image-search?q=${encodeURIComponent(query)}&count=${encodeURIComponent(String(count))}`);
    return result(call, response.ok, await response.text());
  }

  if (call.name === "web_image_import") {
    const imageUrl = typeof call.arguments.imageUrl === "string" ? call.arguments.imageUrl.trim() : "";
    if (!imageUrl) return result(call, false, "web_image_import requires imageUrl.");
    const asset = await imageAssetFromUrl(imageUrl);
    return result(call, true, JSON.stringify({ ok: true, asset, availableImageAssets: listUploadedAssets(), note: "Use asset.id as assetId or asset.name as assetName with Office image tools or generated Office file tools." }));
  }

  return null;
}

async function m365Tool(call: ToolCallRequest) {
  if (call.name === "m365_try_office_sso") {
    const sso = await primeM365OfficeSso();
    return result(call, sso.ok, JSON.stringify({ ok: sso.ok, source: sso.source, message: sso.message || (sso.ok ? "Office SSO succeeded." : "Office SSO unavailable.") }));
  }

  await primeM365OfficeSso().catch(() => undefined);

  if (call.name === "m365_auth_status") {
    const response = await fetch("/api/m365/status");
    return result(call, response.ok, await response.text());
  }

  if (call.name === "m365_search_files") {
    const query = typeof call.arguments.query === "string" ? call.arguments.query.trim() : "";
    const top = typeof call.arguments.top === "number" ? call.arguments.top : 10;
    if (!query) return result(call, false, "m365_search_files requires query.");
    const response = await fetch(`/api/m365/search?q=${encodeURIComponent(query)}&top=${encodeURIComponent(String(top))}`);
    return result(call, response.ok, await response.text());
  }

  if (call.name === "m365_read_file") {
    const id = typeof call.arguments.id === "string" ? call.arguments.id.trim() : "";
    const maxChars = typeof call.arguments.maxChars === "number" ? call.arguments.maxChars : 12000;
    if (!id) return result(call, false, "m365_read_file requires id.");
    const response = await fetch(`/api/m365/read?id=${encodeURIComponent(id)}&maxChars=${encodeURIComponent(String(maxChars))}`);
    return result(call, response.ok, await response.text());
  }

  return null;
}

async function resolveGeneratedImage(image: any) {
  if (!image || typeof image !== "object") return image;
  if (typeof image.base64 === "string" && image.base64.trim()) return image;
  const asset = findUploadedAsset(image.assetId || image.assetName || image.imageUrl) || (typeof image.imageUrl === "string" && image.imageUrl.trim() ? await imageAssetFromUrl(image.imageUrl.trim()) : null);
  if (!asset) return image;
  return {
    ...image,
    name: image.name || asset.name,
    type: image.type || asset.type,
    pixelWidth: image.pixelWidth || asset.pixelWidth,
    pixelHeight: image.pixelHeight || asset.pixelHeight,
    base64: asset.base64,
    altText: image.altText || asset.name,
  };
}

async function resolveGeneratedOfficeArguments(call: ToolCallRequest) {
  if (call.name === "powerpoint_generate_deck_file") {
    const templateLookup = call.arguments.templateAssetId || call.arguments.templateAssetName;
    const templateAsset = findUploadedAsset(templateLookup);
    const slides = Array.isArray(call.arguments.slides) ? call.arguments.slides : [];
    const resolvedSlides = [];
    for (const slide of slides) {
      if (!slide || typeof slide !== "object") {
        resolvedSlides.push(slide);
        continue;
      }
      const images = Array.isArray((slide as any).images) ? (slide as any).images : [];
      const resolvedImages = [];
      for (const image of images) resolvedImages.push(await resolveGeneratedImage(image));
      resolvedSlides.push({ ...slide, images: resolvedImages });
    }
    return { ...call.arguments, ...(templateAsset ? { template: { name: templateAsset.name, type: templateAsset.type, base64: templateAsset.base64, brandProfile: templateAsset.brandProfile } } : {}), slides: resolvedSlides };
  }

  if (call.name === "word_generate_document_file") {
    const templateLookup = call.arguments.templateAssetId || call.arguments.templateAssetName;
    const templateAsset = findUploadedAsset(templateLookup);
    const resolvedTopImages = [];
    for (const image of Array.isArray(call.arguments.images) ? call.arguments.images : []) resolvedTopImages.push(await resolveGeneratedImage(image));
    const resolvedSections = [];
    for (const section of Array.isArray(call.arguments.sections) ? call.arguments.sections : []) {
      if (!section || typeof section !== "object") {
        resolvedSections.push(section);
        continue;
      }
      const resolvedImages = [];
      for (const image of Array.isArray((section as any).images) ? (section as any).images : []) resolvedImages.push(await resolveGeneratedImage(image));
      resolvedSections.push({ ...section, images: resolvedImages });
    }
    return { ...call.arguments, ...(templateAsset ? { templateBase64: templateAsset.base64, templateAssetName: templateAsset.name } : {}), images: resolvedTopImages, ...(Array.isArray(call.arguments.sections) ? { sections: resolvedSections } : {}) };
  }

  if (call.name === "excel_generate_workbook_file") {
    const resolvedSheets = [];
    for (const sheet of Array.isArray(call.arguments.sheets) ? call.arguments.sheets : []) {
      if (!sheet || typeof sheet !== "object") {
        resolvedSheets.push(sheet);
        continue;
      }
      const resolvedImages = [];
      for (const image of Array.isArray((sheet as any).images) ? (sheet as any).images : []) resolvedImages.push(await resolveGeneratedImage(image));
      resolvedSheets.push({ ...sheet, images: resolvedImages });
    }
    return { ...call.arguments, ...(Array.isArray(call.arguments.sheets) ? { sheets: resolvedSheets } : {}) };
  }

  return call.arguments;
}

async function generatedOfficeTool(call: ToolCallRequest) {
  const path = call.name === "powerpoint_generate_deck_file" ? "/api/generated/pptx" : call.name === "word_generate_document_file" ? "/api/generated/docx" : call.name === "excel_generate_workbook_file" ? "/api/generated/xlsx" : null;
  if (!path) return null;
  const response = await fetch(path, {
    method: "POST",
    headers: { "content-type": "application/json", accept: "application/json" },
    body: JSON.stringify(await resolveGeneratedOfficeArguments(call)),
  });
  const kind = call.name === "powerpoint_generate_deck_file" ? "pptx" : call.name === "word_generate_document_file" ? "docx" : "xlsx";
  return result(call, response.ok, withArtifactLink(kind, await response.text()));
}

export function normalizeToolCall(raw: any, index: number): ToolCallRequest {
  return {
    id: raw?.id || `tool-${Date.now()}-${index}`,
    name: raw?.function?.name || raw?.name || "unknown_tool",
    arguments: parseArgs(raw?.function?.arguments ?? raw?.arguments),
  };
}

export async function executeToolCall(host: OfficeHost, call: ToolCallRequest): Promise<ToolCallResult> {
  const demo = await demoShowcaseTool(host, call);
  if (demo) return demo;
  const web = await webTool(call);
  if (web) return web;
  const m365 = await m365Tool(call);
  if (m365) return m365;
  const generated = await generatedOfficeTool(call);
  if (generated) return generated;
  return executeOfficeTool(host, call);
}




