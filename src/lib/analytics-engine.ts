import type { ColumnProfile, DataRecord, DatasetSchema } from "./dataset-schema";

export type AggMode = "count" | "sum" | "avg";

export interface PivotQuery {
  rows: string;
  columns?: string;
  metric?: string;
  agg: AggMode;
}

function num(v: unknown) {
  const n = Number(String(v ?? "").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

export function runPivot(records: DataRecord[], query: PivotQuery) {
  const group = new Map<string, { row: string; col: string; vals: number[]; count: number }>();
  records.forEach((r) => {
    const rk = String(r[query.rows] ?? "Non renseigné");
    const ck = query.columns ? String(r[query.columns] ?? "—") : "Total";
    const key = `${rk}__${ck}`;
    const hit = group.get(key) ?? { row: rk, col: ck, vals: [], count: 0 };
    hit.count += 1;
    if (query.metric) hit.vals.push(num(r[query.metric]));
    group.set(key, hit);
  });
  return [...group.values()].map((g) => {
    const sum = g.vals.reduce((a, b) => a + b, 0);
    const value =
      query.agg === "count"
        ? g.count
        : query.agg === "avg"
          ? (g.vals.length ? sum / g.vals.length : 0)
          : sum;
    return { row: g.row, col: g.col, value };
  });
}

function dimensionScore(c: ColumnProfile, rowCount: number) {
  let s = 0;
  if (c.inferredType === "category" || c.inferredType === "text") s += 50;
  if (c.inferredType === "date") s += 35;
  if (c.uniqueCount <= 100 && c.uniqueCount >= 2) s += 25;
  if (c.uniqueCount > Math.min(400, rowCount * 0.9)) s -= 60;
  return s;
}

/** Always returns at least one pivot when columns exist — avoids empty BI after upload. */
export function suggestPivotQueries(schema: DatasetSchema): PivotQuery[] {
  if (!schema.columns.length) return [];

  const numerics = schema.columns.filter((c) =>
    ["numeric", "currency", "percentage"].includes(c.inferredType),
  );

  const ranked = [...schema.columns]
    .map((c) => ({ c, score: dimensionScore(c, schema.rowCount) }))
    .sort((a, b) => b.score - a.score);

  const rowDim = ranked[0]?.c ?? schema.columns[0];
  const colDim =
    ranked.find((x) => x.c.key !== rowDim.key && x.score > 10)?.c ??
    ranked.find((x) => x.c.key !== rowDim.key)?.c;

  const colsKey = colDim && colDim.key !== rowDim.key ? colDim.key : undefined;

  const raw: PivotQuery[] = [
    { rows: rowDim.key, columns: colsKey, agg: "count" },
  ];

  if (numerics[0]) {
    raw.push({
      rows: rowDim.key,
      columns: colsKey,
      metric: numerics[0].key,
      agg: "sum",
    });
    raw.push({
      rows: rowDim.key,
      columns: colsKey,
      metric: numerics[0].key,
      agg: "avg",
    });
  }

  const altRow = ranked.find((x) => x.c.key !== rowDim.key)?.c;
  if (altRow && raw.length < 6) {
    raw.push({ rows: altRow.key, agg: "count" });
  }

  const seen = new Set<string>();
  const deduped: PivotQuery[] = [];
  for (const q of raw) {
    const k = `${q.rows}|${q.columns ?? ""}|${q.metric ?? ""}|${q.agg}`;
    if (seen.has(k)) continue;
    seen.add(k);
    deduped.push(q);
  }
  return deduped;
}
