export type TableRows = string[][];

function stripCodeFence(text: string) {
  const match = text.match(/```(?:csv|tsv|json|table)?\s*([\s\S]*?)```/i);
  return match?.[1]?.trim() || text.trim();
}

function splitDelimitedLine(line: string, delimiter: string) {
  const values: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"' && next === '"' && inQuotes) {
      current += '"';
      i += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      values.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  values.push(current.trim());
  return values;
}

function normalizeRows(rows: TableRows) {
  const meaningful = rows.filter((row) => row.some((cell) => cell.trim()));
  const width = Math.max(0, ...meaningful.map((row) => row.length));
  return meaningful.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
}

function tryJsonRows(text: string): TableRows | null {
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed) && parsed.length && parsed.every((item) => item && typeof item === "object" && !Array.isArray(item))) {
      const headers = Array.from(new Set(parsed.flatMap((item) => Object.keys(item))));
      return [headers, ...parsed.map((item) => headers.map((header) => String(item[header] ?? "")))];
    }
    if (Array.isArray(parsed) && parsed.every(Array.isArray)) {
      return parsed.map((row) => row.map((value) => String(value ?? "")));
    }
  } catch {
    return null;
  }
  return null;
}

function tryMarkdownTable(text: string): TableRows | null {
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const tableLines = lines.filter((line) => line.includes("|")).map((line) => line.replace(/^\|/, "").replace(/\|$/, ""));
  if (tableLines.length < 2) return null;

  const rows = tableLines
    .filter((line) => !/^\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+$/.test(line))
    .map((line) => line.split("|").map((cell) => cell.trim()));

  const normalized = normalizeRows(rows);
  return normalized.length >= 2 && normalized[0].length >= 2 ? normalized : null;
}

function tryDelimitedRows(text: string): TableRows | null {
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (lines.length < 2) return null;

  const delimiter = lines.some((line) => line.includes("\t")) ? "\t" : lines.some((line) => line.includes(",")) ? "," : null;
  if (!delimiter) return null;

  const rows = normalizeRows(lines.map((line) => splitDelimitedLine(line, delimiter)));
  const hasMultipleColumns = rows.some((row) => row.length > 1);
  return hasMultipleColumns ? rows : null;
}

export function extractTableRows(text: string): TableRows | null {
  const cleaned = stripCodeFence(text);
  return tryJsonRows(cleaned) ?? tryMarkdownTable(cleaned) ?? tryDelimitedRows(cleaned);
}

export function excelInsertionPreview(text: string) {
  const rows = extractTableRows(text);
  if (!rows) return { kind: "single-cell" as const, rows: [[text]] };
  return { kind: "table" as const, rows };
}

export function splitTextForSlides(text: string) {
  const cleaned = text.trim();
  if (!cleaned) return [""];

  const explicit = cleaned.split(/^\s*---\s*(?:slide)?\s*$/gim).map((part) => part.trim()).filter(Boolean);
  if (explicit.length > 1) return explicit;

  const headingSections = cleaned
    .split(/(?=^#{1,2}\s+|^[A-Z][^\n]{3,80}\n[-=]{3,}\s*$)/gm)
    .map((part) => part.trim())
    .filter(Boolean);
  if (headingSections.length > 1) return headingSections;

  const paragraphs = cleaned.split(/\n\s*\n/g).map((part) => part.trim()).filter(Boolean);
  const slides: string[] = [];
  let current = "";
  for (const paragraph of paragraphs) {
    const next = current ? `${current}\n\n${paragraph}` : paragraph;
    if (next.length > 900 && current) {
      slides.push(current);
      current = paragraph;
    } else {
      current = next;
    }
  }
  if (current) slides.push(current);
  return slides.length ? slides : [cleaned];
}

export function summarizeAttachment(name: string, text: string, maxChars = 12000) {
  const trimmed = text.trim();
  const clipped = trimmed.length > maxChars ? `${trimmed.slice(0, maxChars)}\n\n[Attachment clipped after ${maxChars.toLocaleString()} characters.]` : trimmed;
  return [`Attached file: ${name}`, clipped].join("\n");
}