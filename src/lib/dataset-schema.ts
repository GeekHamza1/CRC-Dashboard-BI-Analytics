export type DataValue = string | number | boolean | Date | null | undefined;
export type DataRecord = Record<string, DataValue>;

export type InferredDataType =
  | "text"
  | "numeric"
  | "date"
  | "category"
  | "percentage"
  | "currency";

export interface ColumnProfile {
  key: string;
  label: string;
  nonNullCount: number;
  uniqueCount: number;
  samples: string[];
  inferredType: InferredDataType;
}

export interface DatasetSchema {
  columns: ColumnProfile[];
  rowCount: number;
}

function asString(v: DataValue) {
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function parseNumber(text: string): number | null {
  const clean = text.replace(/\s/g, "").replace(",", ".").replace(/[^\d.-]/g, "");
  if (!clean) return null;
  const n = Number(clean);
  return Number.isFinite(n) ? n : null;
}

function looksLikeDate(text: string) {
  if (!text) return false;
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return true;
  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}/.test(text)) return true;
  return !Number.isNaN(Date.parse(text));
}

function inferType(values: string[]): InferredDataType {
  if (!values.length) return "text";
  const lower = values.map((v) => v.toLowerCase().trim());
  const numericHits = values.filter((v) => parseNumber(v) != null).length;
  const dateHits = values.filter((v) => looksLikeDate(v)).length;
  const pctHits = lower.filter((v) => v.includes("%")).length;
  const currencyHits = lower.filter((v) => /(mad|dh|eur|usd|€|\$)/.test(v)).length;
  if (pctHits / values.length > 0.45) return "percentage";
  if (currencyHits / values.length > 0.35) return "currency";
  if (dateHits / values.length > 0.6) return "date";
  if (numericHits / values.length > 0.65) return "numeric";
  const uniq = new Set(values.map((x) => x.trim().toLowerCase())).size;
  if (uniq <= Math.max(40, Math.round(values.length * 0.15))) return "category";
  return "text";
}

export function discoverDatasetSchema(rows: DataRecord[]): DatasetSchema {
  const keys = [...new Set(rows.flatMap((r) => Object.keys(r)))];
  const columns = keys.map((key) => {
    const vals = rows.map((r) => r[key]);
    const raw = vals
      .map((v) => {
        if (v instanceof Date && !Number.isNaN(+v)) return v.toISOString().slice(0, 10);
        if (typeof v === "number" && Number.isFinite(v)) return String(v);
        if (typeof v === "boolean") return v ? "1" : "0";
        return asString(v).trim();
      })
      .filter(Boolean);
    const unique = new Set(raw);
    return {
      key,
      label: key,
      nonNullCount: vals.filter((v) => v !== null && v !== undefined && v !== "").length,
      uniqueCount: unique.size,
      samples: [...unique].slice(0, 6),
      inferredType: inferType(raw.length ? raw : ["—"]),
    } satisfies ColumnProfile;
  });
  return { columns, rowCount: rows.length };
}

export function schemaStorageKey(datasetId: string) {
  return `bi:schema:${datasetId}`;
}

export function saveSchemaLocal(datasetId: string, schema: DatasetSchema) {
  if (typeof window === "undefined") return;
  localStorage.setItem(schemaStorageKey(datasetId), JSON.stringify(schema));
}

export function loadSchemaLocal(datasetId: string): DatasetSchema | null {
  if (typeof window === "undefined") return null;
  const raw = localStorage.getItem(schemaStorageKey(datasetId));
  if (!raw) return null;
  try {
    return JSON.parse(raw) as DatasetSchema;
  } catch {
    return null;
  }
}
