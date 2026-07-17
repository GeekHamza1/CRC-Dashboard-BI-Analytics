"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import { ThemeToggle } from "@/components/theme-toggle";
import {
  computeNumericStats,
  pickChartDimensions,
  pickDateColumn,
  seriesByDateColumn,
  topValueCounts,
} from "@/lib/bi-auto-analytics";
import { parseBiFile } from "@/lib/bi-parse";
import {
  exportDomAsPng,
  exportRowsExcel,
  exportTablePdf,
  exportTablePptx,
  type TableSlice,
} from "@/lib/bi-export";
import { runPivot, suggestPivotQueries, type PivotQuery } from "@/lib/analytics-engine";
import type { DataRecord } from "@/lib/dataset-schema";
import { buildSchema } from "@/lib/schema-engine";

const PAGE_SIZE = 25;
const PIE_COLORS = ["#f97316", "#3b82f6", "#10b981", "#a855f7", "#eab308", "#ec4899", "#06b6d4", "#64748b"];

function GlassCard({
  title,
  subtitle,
  children,
  action,
}: {
  title: string;
  subtitle?: string;
  children: ReactNode;
  action?: ReactNode;
}) {
  return (
    <section className="glass-panel p-5 sm:p-6 hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-slate-900/10 dark:hover:shadow-slate-950/40 motion-safe:animate-fade-in">
      <div className="flex flex-col gap-1 sm:flex-row sm:items-start sm:justify-between mb-4">
        <div>
          <h2 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50">{title}</h2>
          {subtitle ? <p className="text-sm text-slate-600 dark:text-slate-400 mt-0.5">{subtitle}</p> : null}
        </div>
        {action}
      </div>
      {children}
    </section>
  );
}

function WidgetExportBar(props: {
  disabled?: boolean;
  domRef?: React.RefObject<HTMLElement | null>;
  basename: string;
  title: string;
  subtitle?: string;
  tableColumns?: string[];
  tableRows?: TableSlice[];
}) {
  const [busy, setBusy] = useState(false);
  const { disabled, domRef, basename, title, subtitle, tableColumns = [], tableRows = [] } = props;

  return (
    <div className="flex flex-wrap gap-2 justify-end items-center">
      {busy ? <span className="text-[11px] text-slate-500 animate-pulse">Export…</span> : null}
      <button
        type="button"
        disabled={disabled || busy || !domRef?.current}
        className="rounded-full border border-slate-300 dark:border-slate-600 px-3 py-1.5 text-[11px] font-semibold hover:bg-white/70 dark:hover:bg-slate-800 disabled:opacity-35"
        onClick={async () => {
          if (!domRef?.current) return;
          setBusy(true);
          try {
            await exportDomAsPng(domRef.current as HTMLElement, `${basename}_capture`);
          } catch {
            alert("Export PNG : installez « html-to-image » (npm install) puis rechargez.");
          } finally {
            setBusy(false);
          }
        }}
      >
        PNG
      </button>
      <button
        type="button"
        disabled={disabled || busy || tableRows.length === 0}
        className="rounded-full border border-slate-300 dark:border-slate-600 px-3 py-1.5 text-[11px] font-semibold hover:bg-white/70 dark:hover:bg-slate-800 disabled:opacity-35"
        onClick={async () => {
          setBusy(true);
          try {
            await exportTablePdf(
              title,
              subtitle || "Business Intelligence",
              tableColumns,
              tableRows,
              `${basename}_table`,
            );
          } finally {
            setBusy(false);
          }
        }}
      >
        PDF
      </button>
      <button
        type="button"
        disabled={disabled || busy || tableRows.length === 0}
        className="rounded-full border border-slate-300 dark:border-slate-600 px-3 py-1.5 text-[11px] font-semibold hover:bg-white/70 dark:hover:bg-slate-800 disabled:opacity-35"
        onClick={() => {
          exportRowsExcel(tableColumns, tableRows, title.slice(0, 22), `${basename}_table`);
        }}
      >
        Excel
      </button>
      <button
        type="button"
        disabled={disabled || busy || tableRows.length === 0}
        className="rounded-full border border-slate-300 dark:border-slate-600 px-3 py-1.5 text-[11px] font-semibold hover:bg-white/70 dark:hover:bg-slate-800 disabled:opacity-35"
        onClick={async () => {
          setBusy(true);
          try {
            await exportTablePptx(
              title,
              subtitle || "Synthèse de l'analyse",
              tableColumns,
              tableRows,
              `${basename}_table`,
            );
          } finally {
            setBusy(false);
          }
        }}
      >
        PowerPoint
      </button>
    </div>
  );
}

function formatCell(v: unknown): string {
  if (v == null) return "";
  if (v instanceof Date && !Number.isNaN(+v)) return v.toLocaleString("fr-FR");
  if (typeof v === "boolean") return v ? "oui" : "non";
  return String(v);
}

function compareRaw(a: unknown, b: unknown): number {
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;
  const na = typeof a === "number" ? a : Number(String(a).replace(",", "."));
  const nb = typeof b === "number" ? b : Number(String(b).replace(",", "."));
  if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
  if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime();
  return String(a).localeCompare(String(b), "fr");
}

function pivotQueriesEqual(a: PivotQuery | null | undefined, b: PivotQuery | null | undefined): boolean {
  if (!a || !b) return false;
  return (
    a.rows === b.rows &&
    (a.columns ?? "") === (b.columns ?? "") &&
    (a.metric ?? "") === (b.metric ?? "") &&
    a.agg === b.agg
  );
}

export default function BusinessIntelligencePage() {
  const [rows, setRows] = useState<DataRecord[]>([]);
  const [datasetId, setDatasetId] = useState("bi_workspace");
  const [query, setQuery] = useState<PivotQuery | null>(null);
  const [busy, setBusy] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [fileMeta, setFileMeta] = useState<{ name: string; sheet: string; sheets: string[] } | null>(null);

  const [previewPage, setPreviewPage] = useState(0);
  const [sortKey, setSortKey] = useState<string | null>(null);
  const [sortDir, setSortDir] = useState<"asc" | "desc">("asc");

  const chartBoxRef = useRef<HTMLDivElement>(null);
  const pivotBoxRef = useRef<HTMLDivElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);

  const schema = useMemo(() => (rows.length ? buildSchema(datasetId, rows) : null), [datasetId, rows]);

  const suggestions = useMemo(() => (schema ? suggestPivotQueries(schema) : []), [schema]);

  useEffect(() => {
    setQuery(null);
    setPreviewPage(0);
    setSortKey(null);
    setSortDir("asc");
  }, [rows]);

  const activeQuery = useMemo(() => {
    if (query) return query;
    const d = suggestions[0];
    if (!d) return null;
    return d;
  }, [query, suggestions]);

  const pivot = useMemo(() => (activeQuery && rows.length ? runPivot(rows, activeQuery) : []), [activeQuery, rows]);

  const pivotExport = useMemo<TableSlice[]>(() => {
    return pivot.map((p) => ({
      Rangée: p.row,
      Segment: activeQuery?.columns ? p.col : "Total",
      Mesure: Number(p.value.toFixed(6)),
      Agrégation: activeQuery?.agg ?? "count",
    }));
  }, [activeQuery?.agg, activeQuery?.columns, pivot]);

  const pivotCols = ["Rangée", "Segment", "Mesure", "Agrégation"];

  const chartData = useMemo(() => {
    const grouped = new Map<string, number>();
    pivot.forEach((p) => grouped.set(p.row, (grouped.get(p.row) ?? 0) + p.value));
    return [...grouped.entries()].map(([name, value]) => ({ name, value }));
  }, [pivot]);

  const numericSummaries = useMemo(() => {
    if (!schema) return [];
    return schema.columns
      .map((c) => computeNumericStats(rows, c))
      .filter((x): x is NonNullable<typeof x> => Boolean(x));
  }, [rows, schema]);

  const chartDims = useMemo(() => (schema ? pickChartDimensions(schema) : []), [schema]);

  const barCharts = useMemo(() => {
    if (!schema || !chartDims.length) return [];
    return chartDims.map((col) => ({
      key: col.key,
      title: col.label,
      subtitle: `${col.inferredType} · ${col.uniqueCount} valeurs distinctes`,
      data: topValueCounts(rows, col.key, 16),
    }));
  }, [chartDims, rows, schema]);

  const dateCol = useMemo(() => (schema ? pickDateColumn(schema) : null), [schema]);

  const dateSeries = useMemo(() => {
    if (!dateCol || !rows.length) return [];
    return seriesByDateColumn(rows, dateCol.key).slice(-90);
  }, [dateCol, rows]);

  const pieCharts = useMemo(() => {
    return barCharts.filter((b) => b.data.length <= 12 && b.data.length >= 2).slice(0, 3);
  }, [barCharts]);

  const previewColumns = useMemo(() => schema?.columns.map((c) => c.key) ?? [], [schema]);

  const sortedPreviewRows = useMemo(() => {
    const list = [...rows];
    if (sortKey) {
      const dir = sortDir === "asc" ? 1 : -1;
      list.sort((a, b) => compareRaw(a[sortKey], b[sortKey]) * dir);
    }
    return list;
  }, [rows, sortDir, sortKey]);

  const previewSlice = useMemo(() => {
    const start = previewPage * PAGE_SIZE;
    return sortedPreviewRows.slice(start, start + PAGE_SIZE);
  }, [previewPage, sortedPreviewRows]);

  const totalPages = Math.max(1, Math.ceil(sortedPreviewRows.length / PAGE_SIZE));

  const loadFile = useCallback(async (file: File | undefined) => {
    if (!file) return;
    setBusy(true);
    setParseError(null);
    try {
      const parsed = await parseBiFile(file);
      if (!parsed.rows.length) {
        setParseError(
          "Aucune ligne exploitable trouvée. Vérifiez que la première feuille contient des en-têtes et des données, ou essayez un autre onglet exporté en CSV.",
        );
        setRows([]);
        setFileMeta(null);
        return;
      }
      setRows(parsed.rows);
      setDatasetId(`bi_${file.name.replace(/\W+/g, "_").slice(0, 48)}_${parsed.rows.length}`);
      setFileMeta({
        name: file.name,
        sheet: parsed.sheetName,
        sheets: parsed.workbookSheets,
      });
    } catch (err) {
      console.error(err);
      setParseError(err instanceof Error ? err.message : "Lecture du fichier impossible.");
      setRows([]);
      setFileMeta(null);
    } finally {
      setBusy(false);
    }
  }, []);

  const toggleSort = (key: string) => {
    if (sortKey !== key) {
      setSortKey(key);
      setSortDir("asc");
      return;
    }
    setSortDir((d) => (d === "asc" ? "desc" : "asc"));
  };

  const kpiCards = useMemo(() => {
    if (!schema) return [];
    const cards: { title: string; subtitle?: string; value: ReactNode }[] = [
      {
        title: "Lignes chargées",
        subtitle: fileMeta?.name ? `Fichier : ${fileMeta.name}` : undefined,
        value: schema.rowCount.toLocaleString("fr-FR"),
      },
      {
        title: "Colonnes détectées",
        subtitle: fileMeta?.sheet ? `Feuille : ${fileMeta.sheet}` : undefined,
        value: schema.columns.length,
      },
      {
        title: "Dimensions catégorielles",
        subtitle: "Colonnes texte / catégorie",
        value: schema.columns.filter((c) => c.inferredType === "category" || c.inferredType === "text").length,
      },
      {
        title: "Champs numériques",
        subtitle: "Sommes & moyennes auto ci-dessous",
        value: schema.columns.filter((c) =>
          ["numeric", "currency", "percentage"].includes(c.inferredType),
        ).length,
      },
    ];

    const topDistinct = [...schema.columns].sort((a, b) => b.uniqueCount - a.uniqueCount)[0];
    if (topDistinct) {
      cards.push({
        title: "Cardinalité max",
        subtitle: topDistinct.key,
        value: topDistinct.uniqueCount.toLocaleString("fr-FR"),
      });
    }

    return cards;
  }, [fileMeta?.name, fileMeta?.sheet, schema]);

  return (
    <div className="max-w-[1640px] mx-auto px-4 pb-24 pt-10 space-y-8">
      <header className="flex flex-wrap items-start gap-4 justify-between">
        <div>
          <p className="text-[11px] uppercase tracking-[0.22em] text-sky-600 dark:text-sky-300 font-bold">
            Analyse universelle
          </p>
          <h1 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white tracking-tight">
            Intelligence décisionnelle multi-fichiers
          </h1>
          <p className="text-sm text-slate-600 dark:text-slate-400 max-w-2xl mt-2 leading-relaxed">
            Import Excel / CSV entièrement côté navigateur : inférence de schéma, KPI, graphiques automatiques,
            tableau paginé et suggestions d'analyse — même design que le tableau CRC.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <Link href="/" className="rounded-full px-3 py-2 text-xs font-semibold border border-slate-300 dark:border-slate-600">
            Tableau CRC
          </Link>
          <ThemeToggle />
        </div>
      </header>

      <GlassCard
        title="1. Importer un jeu de données"
        subtitle="Formats : .xlsx, .xls, .csv — la première feuille non vide est analysée automatiquement."
        action={
          <label className="rounded-full px-5 py-2 text-xs font-bold bg-gradient-to-r from-sky-500 to-indigo-500 text-white cursor-pointer shadow disabled:opacity-50">
            {busy ? "Analyse…" : "Choisir un fichier"}
            <input
              type="file"
              className="hidden"
              disabled={busy}
              accept=".csv,.xls,.xlsx,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={(e) => {
                void loadFile(e.target.files?.[0]);
                e.target.value = "";
              }}
            />
          </label>
        }
      >
        <div
          className="rounded-2xl border-2 border-dashed border-slate-300 dark:border-slate-600 px-4 py-8 text-center bg-white/40 dark:bg-slate-950/30"
          onDragOver={(ev) => {
            ev.preventDefault();
          }}
          onDrop={(ev) => {
            ev.preventDefault();
            const f = ev.dataTransfer.files?.[0];
            if (f) void loadFile(f);
          }}
        >
          {busy ? (
            <p className="text-sm text-sky-700 dark:text-sky-300 animate-pulse">
              Lecture du classeur, détection des colonnes…
            </p>
          ) : parseError ? (
            <p className="text-sm text-rose-700 dark:text-rose-300">{parseError}</p>
          ) : fileMeta ? (
            <div className="space-y-1 text-sm text-slate-700 dark:text-slate-200">
              <p className="font-semibold">{fileMeta.name}</p>
              <p className="text-xs text-slate-500">
                Feuille active : <strong>{fileMeta.sheet}</strong>
                {fileMeta.sheets.length > 1 ? ` · ${fileMeta.sheets.length} feuilles dans le classeur` : null}
              </p>
            </div>
          ) : (
            <p className="text-sm text-slate-600 dark:text-slate-400">
              Glissez-déposez un fichier ou utilisez le bouton « Choisir un fichier ».
            </p>
          )}
        </div>
      </GlassCard>

      {schema && rows.length > 0 && (
        <>
          <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-5">
            {kpiCards.map((k) => (
              <GlassCard key={k.title} title={k.title} subtitle={k.subtitle}>
                <p className="text-3xl font-bold text-indigo-600 dark:text-indigo-300 tabular-nums mt-1">{k.value}</p>
              </GlassCard>
            ))}
          </div>

          {numericSummaries.length > 0 ? (
            <GlassCard
              title="Synthèse numérique automatique"
              subtitle="Somme, moyenne, min, max sur les colonnes détectées comme numériques."
            >
              <div className="overflow-x-auto rounded-2xl border border-slate-200/70 dark:border-slate-700/80">
                <table className="min-w-[720px] w-full text-xs">
                  <thead className="bg-slate-900 text-white">
                    <tr>
                      <th className="px-3 py-2 text-left">Colonne</th>
                      <th className="px-3 py-2 text-right">Somme</th>
                      <th className="px-3 py-2 text-right">Moyenne</th>
                      <th className="px-3 py-2 text-right">Min</th>
                      <th className="px-3 py-2 text-right">Max</th>
                      <th className="px-3 py-2 text-right">Occurrences numériques</th>
                    </tr>
                  </thead>
                  <tbody>
                    {numericSummaries.map((s, idx) => (
                      <tr
                        key={s.key}
                        className={idx % 2 === 0 ? "bg-white/90 dark:bg-slate-950/40" : "bg-slate-50 dark:bg-slate-900/45"}
                      >
                        <td className="px-3 py-1.5 font-medium">{s.key}</td>
                        <td className="px-3 py-1.5 text-right tabular-nums">{s.sum.toLocaleString("fr-FR", { maximumFractionDigits: 2 })}</td>
                        <td className="px-3 py-1.5 text-right tabular-nums">{s.avg.toLocaleString("fr-FR", { maximumFractionDigits: 2 })}</td>
                        <td className="px-3 py-1.5 text-right tabular-nums">{s.min.toLocaleString("fr-FR", { maximumFractionDigits: 2 })}</td>
                        <td className="px-3 py-1.5 text-right tabular-nums">{s.max.toLocaleString("fr-FR", { maximumFractionDigits: 2 })}</td>
                        <td className="px-3 py-1.5 text-right tabular-nums">{s.count}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </GlassCard>
          ) : null}

          <GlassCard
            title="Découverte de schéma"
            subtitle="Types inférés à partir des valeurs : texte, numérique, date, catégorie, pourcentage, devise."
          >
            <div className="overflow-auto max-h-72 rounded-2xl border border-slate-200/70 dark:border-slate-700/80">
              <table className="min-w-full text-xs">
                <thead className="bg-slate-900 text-white sticky top-0 z-10">
                  <tr>
                    <th className="px-3 py-2 text-left">Colonne</th>
                    <th className="px-3 py-2 text-left">Type</th>
                    <th className="px-3 py-2 text-right">Non vides</th>
                    <th className="px-3 py-2 text-right">Distincts</th>
                    <th className="px-3 py-2 text-left">Échantillons</th>
                  </tr>
                </thead>
                <tbody>
                  {schema.columns.map((c) => (
                    <tr key={c.key} className="border-t border-slate-100 dark:border-slate-800">
                      <td className="px-3 py-1.5 font-medium">{c.key}</td>
                      <td className="px-3 py-1.5">{c.inferredType}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{c.nonNullCount}</td>
                      <td className="px-3 py-1.5 text-right tabular-nums">{c.uniqueCount}</td>
                      <td className="px-3 py-1.5 text-slate-600 dark:text-slate-400 max-w-md truncate">
                        {c.samples.join(" · ")}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </GlassCard>

          <GlassCard
            title="Graphiques automatiques"
            subtitle="Histogrammes sur les dimensions les plus lisibles ; secteurs si peu de modalités ; tendance temporelle si une colonne date est détectée."
          >
            <div className="space-y-10">
              {barCharts.length === 0 ? (
                <p className="text-sm text-slate-500">
                    Pas assez de signal catégoriel pour tracer des histogrammes automatiques — utilisez l'analyse ou des
                  colonnes avec moins de valeurs distinctes.
                </p>
              ) : null}
              {barCharts.map((bc) => (
                <div key={bc.key}>
                  <p className="text-sm font-semibold text-slate-800 dark:text-slate-100 mb-1">{bc.title}</p>
                  <p className="text-xs text-slate-500 mb-3">{bc.subtitle}</p>
                  <div className="h-72 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={bc.data} margin={{ bottom: bc.data.length > 10 ? 48 : 16 }}>
                        <CartesianGrid strokeDasharray="3 3" className="stroke-slate-200 dark:stroke-slate-700" />
                        <XAxis
                          dataKey="name"
                          tick={{ fontSize: 10 }}
                          interval={0}
                          angle={bc.data.length > 8 ? -24 : 0}
                          textAnchor={bc.data.length > 8 ? "end" : "middle"}
                          height={bc.data.length > 8 ? 56 : 28}
                        />
                        <YAxis tick={{ fontSize: 11 }} />
                        <Tooltip />
                        <Bar dataKey="value" fill="#f97316" radius={[4, 4, 0, 0]} maxBarSize={42} name="Effectif" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              ))}

              {pieCharts.length > 0 ? (
                <div className="grid lg:grid-cols-3 gap-6">
                  {pieCharts.map((pc) => (
                    <div key={`pie-${pc.key}`}>
                      <p className="text-sm font-semibold mb-2">{pc.title}</p>
                      <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie data={pc.data} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={72} label>
                              {pc.data.map((_, i) => (
                                <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                              ))}
                            </Pie>
                            <Tooltip />
                            <Legend layout="horizontal" verticalAlign="bottom" />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  ))}
                </div>
              ) : null}

              {dateSeries.length > 0 && dateCol ? (
                <div>
                  <p className="text-sm font-semibold text-slate-800 dark:text-slate-100 mb-1">
                    Tendance — {dateCol.key}
                  </p>
                  <p className="text-xs text-slate-500 mb-3">Nombre d&apos;enregistrements par jour (90 derniers points).</p>
                  <div className="h-72 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={dateSeries} margin={{ left: 4, right: 8 }}>
                        <CartesianGrid strokeDasharray="3 3" className="stroke-slate-200 dark:stroke-slate-700" />
                        <XAxis dataKey="period" tick={{ fontSize: 10 }} />
                        <YAxis tick={{ fontSize: 11 }} />
                        <Tooltip />
                        <Legend />
                        <Line type="monotone" dataKey="count" name="Volume" stroke="#3b82f6" strokeWidth={2} dot={false} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              ) : null}
            </div>
          </GlassCard>

          <GlassCard
            title="Aperçu des données"
            subtitle={`${sortedPreviewRows.length.toLocaleString("fr-FR")} lignes · tri par colonne · ${PAGE_SIZE} lignes par page`}
            action={
              <WidgetExportBar
                disabled={!sortedPreviewRows.length}
                domRef={previewRef}
                basename="data_preview"
                title="Aperçu des données"
                subtitle={fileMeta?.name}
                tableColumns={previewColumns}
                tableRows={sortedPreviewRows.slice(0, 500).map((r) =>
                  Object.fromEntries(previewColumns.map((k) => [k, formatCell(r[k])])),
                )}
              />
            }
          >
            <div ref={previewRef} className="overflow-auto max-h-[480px] rounded-2xl border border-slate-200 dark:border-slate-700 shadow-inner">
              <table className="min-w-max w-full text-[11px]">
                <thead className="bg-slate-900 text-white sticky top-0 z-10">
                  <tr>
                    {previewColumns.map((col) => (
                      <th key={col} className="px-2 py-2 text-left whitespace-nowrap">
                        <button
                          type="button"
                          className="inline-flex items-center gap-1 font-semibold hover:text-sky-200"
                          onClick={() => toggleSort(col)}
                        >
                          {col}
                          {sortKey === col ? (sortDir === "asc" ? " ↑" : " ↓") : ""}
                        </button>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewSlice.map((r, i) => (
                    <tr
                      key={`${previewPage}-${i}`}
                      className={i % 2 === 0 ? "bg-white/90 dark:bg-slate-950/35" : "bg-slate-50 dark:bg-slate-900/40"}
                    >
                      {previewColumns.map((col) => (
                        <td key={col} className="px-2 py-1 border-t border-slate-100 dark:border-slate-800 max-w-[220px] truncate">
                          {formatCell(r[col])}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="flex flex-wrap items-center justify-between gap-3 mt-4 text-xs">
              <span className="text-slate-500">
                Page {previewPage + 1} / {totalPages}
              </span>
              <div className="flex gap-2">
                <button
                  type="button"
                  disabled={previewPage <= 0}
                  className="rounded-xl border px-3 py-1.5 dark:border-slate-600 disabled:opacity-35"
                  onClick={() => setPreviewPage((p) => Math.max(0, p - 1))}
                >
                  Précédent
                </button>
                <button
                  type="button"
                  disabled={previewPage >= totalPages - 1}
                  className="rounded-xl border px-3 py-1.5 dark:border-slate-600 disabled:opacity-35"
                  onClick={() => setPreviewPage((p) => Math.min(totalPages - 1, p + 1))}
                >
                  Suivant
                </button>
              </div>
            </div>
          </GlassCard>

          <GlassCard
            title="Suggestions d'analyses"
            subtitle="Cliquez pour appliquer une analyse pré-paramétrée (COUNT / SUM / AVG)."
          >
            <div className="flex flex-wrap gap-2">
              {suggestions.map((s, idx) => (
                <button
                  key={`${s.rows}-${s.metric ?? "x"}-${s.agg}-${idx}`}
                  type="button"
                  onClick={() => setQuery(s)}
                  className={`rounded-full px-4 py-2 text-xs font-semibold border transition ${
                    pivotQueriesEqual(query ?? suggestions[0] ?? null, s)
                      ? "border-orange-500 bg-orange-500/15 text-orange-900 dark:text-orange-100"
                      : "border-slate-300 dark:border-slate-600 hover:bg-white/60 dark:hover:bg-slate-800"
                  }`}
                >
                  #{idx + 1} · {s.rows}
                  {s.columns ? ` × ${s.columns}` : ""} · {s.agg.toUpperCase()}
                  {s.metric ? `(${s.metric})` : ""}
                </button>
              ))}
            </div>
          </GlassCard>

          <GlassCard
            title="Créateur d'analyse"
            subtitle="Affinez dimensions, métrique et agrégation — le graphique principal se met à jour."
            action={
              <WidgetExportBar
                disabled={!pivot.length}
                domRef={pivotBoxRef}
                basename="bi_pivot"
                title="Analyse"
                subtitle={
                  activeQuery
                    ? `${activeQuery.rows}${activeQuery.columns ? ` × ${activeQuery.columns}` : ""} · ${activeQuery.agg}`
                    : ""
                }
                tableColumns={pivotCols}
                tableRows={pivotExport}
              />
            }
          >
            <div className="grid md:grid-cols-4 gap-3 text-sm mb-5">
              <select
                className="rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/85 dark:bg-slate-950/60 px-3 py-2"
                value={activeQuery?.rows ?? ""}
                onChange={(e) =>
                  setQuery((q) => ({
                    ...(q ?? suggestions[0] ?? { agg: "count", rows: schema.columns[0].key }),
                    rows: e.target.value,
                  }))
                }
              >
                {schema.columns.map((c) => (
                  <option key={c.key} value={c.key}>
                    Lignes : {c.key}
                  </option>
                ))}
              </select>
              <select
                className="rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/85 dark:bg-slate-950/60 px-3 py-2"
                value={activeQuery?.columns ?? ""}
                onChange={(e) =>
                  setQuery((q) => ({
                    ...(q ?? suggestions[0] ?? { agg: "count", rows: schema.columns[0].key }),
                    columns: e.target.value || undefined,
                  }))
                }
              >
                <option value="">Colonnes : total simple</option>
                {schema.columns.map((c) => (
                  <option key={c.key} value={c.key}>
                    Colonnes : {c.key}
                  </option>
                ))}
              </select>
              <select
                className="rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/85 dark:bg-slate-950/60 px-3 py-2"
                value={activeQuery?.metric ?? ""}
                onChange={(e) =>
                  setQuery((q) => ({
                    ...(q ?? suggestions[0] ?? { agg: "count", rows: schema.columns[0].key }),
                    metric: e.target.value || undefined,
                  }))
                }
              >
                <option value="">Sans métrique (comptages)</option>
                {schema.columns.map((c) => (
                  <option key={`m-${c.key}`} value={c.key}>
                    Métrique : {c.key}
                  </option>
                ))}
              </select>
              <select
                className="rounded-2xl border border-slate-200 dark:border-slate-700 bg-white/85 dark:bg-slate-950/60 px-3 py-2"
                value={activeQuery?.agg ?? "count"}
                onChange={(e) =>
                  setQuery((q) => ({
                    ...(q ?? suggestions[0] ?? { rows: schema.columns[0].key, agg: "count" }),
                    agg: e.target.value as PivotQuery["agg"],
                  }))
                }
              >
                <option value="count">COUNT</option>
                <option value="sum">SUM</option>
                <option value="avg">AVG</option>
              </select>
            </div>

            <div ref={pivotBoxRef} className="overflow-x-auto rounded-2xl border border-slate-200/70 dark:border-slate-700/70 mb-8">
              <table className="min-w-[640px] w-full text-xs sm:text-sm">
                <thead className="bg-slate-900 text-white">
                  <tr>
                    {pivotCols.map((c) => (
                      <th key={c} className="px-3 py-2 text-left font-semibold">
                        {c}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pivot.slice(0, 400).map((line, idx) => (
                    <tr
                      key={`${line.row}-${line.col}-${idx}`}
                      className={idx % 2 === 0 ? "bg-white/90 dark:bg-slate-950/40" : "bg-slate-50 dark:bg-slate-900/50"}
                    >
                      <td className="px-3 py-1.5">{line.row}</td>
                      <td className="px-3 py-1.5">{activeQuery?.columns ? line.col : "Total"}</td>
                      <td className="px-3 py-1.5 tabular-nums font-semibold text-sky-700 dark:text-sky-300">{line.value}</td>
                      <td className="px-3 py-1.5 uppercase text-[11px] text-slate-500">{activeQuery?.agg ?? "count"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {pivot.length > 400 ? (
                <p className="text-center text-xs text-slate-500 py-2">Affichage limité à 400 lignes — export Excel pour la suite.</p>
              ) : null}
            </div>

            <div className="mt-8 pt-6 border-t border-slate-200/80 dark:border-slate-700/80">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between mb-4">
                <div>
                  <h3 className="text-base font-semibold text-slate-900 dark:text-slate-50">Visualisation du pivot actif</h3>
                  <p className="text-sm text-slate-600 dark:text-slate-400 mt-0.5">
                    Histogramme des rangées (agrégation des segments si colonnes croisées).
                  </p>
                </div>
                <WidgetExportBar
                  disabled={!chartData.length}
                  domRef={chartBoxRef}
                  basename="bi_chart"
                  title="Graphique d'analyse"
                  subtitle={`Mesure : ${activeQuery?.agg ?? "count"}`}
                  tableColumns={["Libellé", "Valeur"]}
                  tableRows={chartData.slice(0, 100).map((d) => ({ Libellé: d.name, Valeur: d.value }))}
                />
              </div>
              <div ref={chartBoxRef} className="h-80 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData.slice(0, 28)} margin={{ left: 4, right: 8, top: 8, bottom: chartData.length > 14 ? 52 : 12 }}>
                    <CartesianGrid strokeDasharray="3 3" className="stroke-slate-200 dark:stroke-slate-700" />
                    <XAxis
                      dataKey="name"
                      tick={{ fontSize: 10 }}
                      interval={0}
                      angle={chartData.length > 12 ? -28 : 0}
                      textAnchor={chartData.length > 12 ? "end" : "middle"}
                      height={chartData.length > 12 ? 70 : 30}
                    />
                    <YAxis tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="value" name="Mesure" fill="#f97316" radius={[6, 6, 2, 2]} maxBarSize={48} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </GlassCard>
        </>
      )}
    </div>
  );
}
