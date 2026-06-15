import * as XLSX from "xlsx";
import type { DataRecord } from "./dataset-schema";

export interface ParseResult {
  rows: DataRecord[];
  sheetName: string;
  workbookSheets: string[];
}

/** Normalize keys from Excel (trim, collapse spaces). */
function normalizeRecord(rec: DataRecord): DataRecord {
  const out: DataRecord = {};
  for (const [k, v] of Object.entries(rec)) {
    const nk = String(k ?? "").trim().replace(/\s+/g, " ");
    if (!nk) continue;
    out[nk] = v;
  }
  return out;
}

function omitCellA1(ws: XLSX.WorkSheet): XLSX.WorkSheet {
  const clone = { ...ws } as XLSX.WorkSheet;
  delete (clone as Record<string, unknown>).A1;
  return clone;
}

function rowsFromSheet(ws: XLSX.WorkSheet): DataRecord[] {
  const raw = XLSX.utils.sheet_to_json<DataRecord>(omitCellA1(ws), {
    defval: null,
    raw: false,
    blankrows: false,
  });
  return raw.map(normalizeRecord).filter((r) => Object.keys(r).length > 0);
}

/**
 * Reads first worksheet that yields at least one row.
 * Falls back to sheet 1 even if empty (caller shows empty state).
 */
export function parseWorkbookBuffer(buf: ArrayBuffer): ParseResult {
  const wb = XLSX.read(buf, {
    type: "array",
    cellDates: true,
    cellNF: false,
    dense: false,
  });
  const names = wb.SheetNames ?? [];
  let sheetName = names[0] ?? "Sheet1";
  let rows: DataRecord[] = [];

  for (const n of names) {
    const ws = wb.Sheets[n];
    if (!ws) continue;
    const candidate = rowsFromSheet(ws);
    if (candidate.length > 0) {
      rows = candidate;
      sheetName = n;
      break;
    }
  }

  if (!rows.length && names.length && wb.Sheets[names[0]]) {
    rows = rowsFromSheet(wb.Sheets[names[0]]);
    sheetName = names[0];
  }

  return { rows, sheetName, workbookSheets: names };
}

/** CSV text → workbook → same pipeline */
export function parseCsvText(text: string): ParseResult {
  const wb = XLSX.read(text, { type: "string", raw: false, cellDates: true });
  const names = wb.SheetNames ?? [];
  const sheetName = names[0] ?? "CSV";
  const ws = wb.Sheets[sheetName];
  const rows = ws ? rowsFromSheet(ws) : [];
  return { rows, sheetName, workbookSheets: names };
}

export async function parseBiFile(file: File): Promise<ParseResult> {
  const name = file.name.toLowerCase();
  const isCsv = name.endsWith(".csv") || file.type === "text/csv";

  if (isCsv) {
    let text = await file.text();
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    return parseCsvText(text);
  }

  const buf = await file.arrayBuffer();
  return parseWorkbookBuffer(buf);
}
