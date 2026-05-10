import type { ColumnProfile, DataRecord, DatasetSchema } from "./dataset-schema";

export interface ValueCount {
  name: string;
  value: number;
}

export interface DateSeriesPoint {
  period: string;
  count: number;
}

export interface NumericColumnStats {
  key: string;
  sum: number;
  avg: number;
  min: number;
  max: number;
  count: number;
}

function num(v: unknown): number | null {
  if (v == null) return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s = String(v).replace(/\s/g, "").replace(",", ".").replace(/[^\d.-]/g, "");
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export function computeNumericStats(rows: DataRecord[], col: ColumnProfile): NumericColumnStats | null {
  if (!["numeric", "currency", "percentage"].includes(col.inferredType)) return null;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  let count = 0;
  for (const r of rows) {
    const n = num(r[col.key]);
    if (n == null) continue;
    sum += n;
    min = Math.min(min, n);
    max = Math.max(max, n);
    count += 1;
  }
  if (!count) return null;
  return {
    key: col.key,
    sum,
    avg: sum / count,
    min,
    max,
    count,
  };
}

export function topValueCounts(rows: DataRecord[], key: string, limit = 15): ValueCount[] {
  const m = new Map<string, number>();
  for (const r of rows) {
    const raw = r[key];
    let label: string;
    if (raw instanceof Date && !Number.isNaN(+raw)) {
      label = raw.toISOString().slice(0, 10);
    } else if (raw == null || raw === "") {
      label = "—";
    } else {
      label = String(raw).trim().slice(0, 120) || "—";
    }
    m.set(label, (m.get(label) ?? 0) + 1);
  }
  return [...m.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([name, value]) => ({ name, value }));
}

export function seriesByDateColumn(rows: DataRecord[], key: string): DateSeriesPoint[] {
  const m = new Map<string, number>();
  for (const r of rows) {
    const raw = r[key];
    let day: string;
    if (raw instanceof Date && !Number.isNaN(+raw)) {
      day = raw.toISOString().slice(0, 10);
    } else {
      const s = String(raw ?? "").trim();
      const d = new Date(s);
      day = Number.isNaN(+d) ? (s.slice(0, 10) || "—") : d.toISOString().slice(0, 10);
    }
    m.set(day, (m.get(day) ?? 0) + 1);
  }
  return [...m.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([period, count]) => ({ period, count }));
}

/** Score columns for automatic bar/pie charts (prefer low-cardinality categories). */
export function pickChartDimensions(schema: DatasetSchema): ColumnProfile[] {
  const { columns, rowCount } = schema;
  const scored = columns
    .map((c) => {
      let score = 0;
      if (c.inferredType === "category") score += 80;
      else if (c.inferredType === "text" && c.uniqueCount <= 80) score += 60;
      else if (c.inferredType === "text") score += 20;
      else if (c.inferredType === "date") score += 40;
      else if (c.inferredType === "numeric") score += 15;
      if (c.uniqueCount <= 12 && c.uniqueCount >= 2) score += 30;
      if (c.uniqueCount > Math.min(500, rowCount)) score -= 40;
      return { c, score };
    })
    .sort((a, b) => b.score - a.score);

  const picked = scored.filter((x) => x.score > 15).slice(0, 4).map((x) => x.c);
  if (picked.length > 0) return picked;
  return scored.slice(0, Math.min(4, scored.length)).map((x) => x.c);
}

export function pickDateColumn(schema: DatasetSchema): ColumnProfile | null {
  const d = schema.columns.find((c) => c.inferredType === "date");
  if (d) return d;
  return schema.columns.find((c) => c.key.toLowerCase().includes("date")) ?? null;
}
