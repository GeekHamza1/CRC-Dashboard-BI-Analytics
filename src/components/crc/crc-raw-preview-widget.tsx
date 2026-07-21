"use client";

import { useMemo, useRef, useState, type MutableRefObject, type ReactNode, type RefObject } from "react";
import {
  Area,
  AreaChart,
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

import { REGION_COLORS, REGION_ORDER, REGION_SHORT } from "@/lib/crc-constants";
import {
  captureElementPng,
  exportRawPreviewExcel,
  exportRawPreviewPdf,
  exportRawPreviewPptx,
} from "@/lib/crc-export-engine";
import type { RawColumnKey } from "@/lib/crc-report-config";
import type { CrcRow } from "@/lib/crc-types";

import type { PivotChartKind } from "./crc-region-pivot-widget";

function GlassShell({
  title,
  subtitle,
  toolbar,
  children,
}: {
  title: string;
  subtitle?: string;
  toolbar: ReactNode;
  children: ReactNode;
}) {
  return (
    <section className="glass-panel p-5 sm:p-6 hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-slate-900/10 dark:hover:shadow-slate-950/40 motion-safe:animate-fade-in">
      <div className="mb-2">
        <h3 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50">{title}</h3>
        {subtitle ? <p className="text-sm text-slate-600 dark:text-slate-400 mt-0.5">{subtitle}</p> : null}
      </div>
      {toolbar}
      {children}
    </section>
  );
}

export function CrcRawPreviewWidget(props: {
  rows: CrcRow[];
  rawColumns: Record<RawColumnKey, boolean>;
  onToggleColumn: (key: RawColumnKey) => void;
  columnCells: {
    key: RawColumnKey;
    label: string;
    tdClass?: string;
    cell: (r: CrcRow) => ReactNode;
  }[];
  isDark: boolean;
  palette: { fg: string; muted: string; grid: string; tooltipBg: string };
  exportBasename: string;
  exportTableRef?: RefObject<HTMLDivElement | null>;
  exportChartRef?: RefObject<HTMLDivElement | null>;
}) {
  const {
    rows,
    rawColumns,
    onToggleColumn,
    columnCells,
    isDark,
    palette,
    exportBasename,
    exportTableRef,
    exportChartRef,
  } = props;

  const visibleDefs = useMemo(
    () => columnCells.filter((c) => rawColumns[c.key]),
    [columnCells, rawColumns],
  );

  const keys = useMemo(() => visibleDefs.map((c) => c.key), [visibleDefs]);
  const headers = useMemo(() => visibleDefs.map((c) => c.label), [visibleDefs]);

  const [colsOpen, setColsOpen] = useState(false);
  const [chartOpen, setChartOpen] = useState(false);
  const [chartKind, setChartKind] = useState<PivotChartKind>("bar");
  const [rowsLimit, setRowsLimit] = useState<number | "ALL">(200);
  const [dateMin, setDateMin] = useState("");
  const [dateMax, setDateMax] = useState("");
  const [query, setQuery] = useState("");
  const chartRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  const filteredRows = useMemo(() => {
    const q = query.trim().toLowerCase();
    return rows.filter((r) => {
      const ymd = r.date
        ? `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}-${String(r.date.getDate()).padStart(2, "0")}`
        : "";
      if (dateMin && (!ymd || ymd < dateMin)) return false;
      if (dateMax && (!ymd || ymd > dateMax)) return false;
      if (!q) return true;
      return [
        r.résultat,
        r.résultatRaw,
        r.téléopérateur,
        r.metier,
        r.natureRéclamation,
        r.regions,
        r.régionCanon,
        r.téléphone,
      ]
        .join(" ")
        .toLowerCase()
        .includes(q);
    });
  }, [rows, dateMin, dateMax, query]);

  const visibleRows = useMemo(
    () => (rowsLimit === "ALL" ? filteredRows : filteredRows.slice(0, rowsLimit)),
    [filteredRows, rowsLimit],
  );

  const regionChartRows = useMemo(
    () =>
      REGION_ORDER.filter((rg) => rows.some((r) => r.régionCanon === rg)).map((rg) => ({
        label: REGION_SHORT[rg],
        value: rows.filter((r) => r.régionCanon === rg).length,
        fill: REGION_COLORS[rg],
      })),
    [rows],
  );

  const pieRegion = useMemo(
    () =>
      REGION_ORDER.filter((rg) => rows.some((r) => r.régionCanon === rg)).map((rg) => ({
        name: REGION_SHORT[rg],
        value: rows.filter((r) => r.régionCanon === rg).length,
        fill: REGION_COLORS[rg],
      })),
    [rows],
  );

  const toolbar = (
    <div className="flex flex-wrap items-center gap-2 mb-4 pb-3 border-b border-slate-200/70 dark:border-slate-700/60">
      <div className="relative">
        <button
          type="button"
          onClick={() => setColsOpen((o) => !o)}
          className="rounded-full px-4 py-2 text-xs font-semibold bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-600"
        >
          Colonnes
        </button>
        {colsOpen ? (
          <div className="absolute left-0 z-30 mt-2 w-64 rounded-2xl border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-900 shadow-xl p-3 space-y-2 max-h-56 overflow-y-auto">
            {columnCells.map((c) => (
              <label key={c.key} className="flex items-center gap-2 text-xs cursor-pointer">
                <input
                  type="checkbox"
                  checked={rawColumns[c.key]}
                  onChange={() => onToggleColumn(c.key)}
                />
                <span>{c.label}</span>
              </label>
            ))}
          </div>
        ) : null}
      </div>
      <button
        type="button"
        onClick={() => setChartOpen((x) => !x)}
        className={`rounded-full px-4 py-2 text-xs font-semibold border ${
          chartOpen ? "border-orange-500 bg-orange-500/15" : "border-slate-200 dark:border-slate-600"
        }`}
      >
        Charts
      </button>
      <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 px-1">Exporter</span>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={async () => {
          const el = chartOpen && chartRef.current ? chartRef.current : tableRef.current;
          if (!el) return;
          await captureElementPng(el, `${exportBasename}_raw`);
        }}
      >
        PNG
      </button>
      <button
        type="button"
        disabled={!keys.length}
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600 disabled:opacity-40"
        onClick={() => void exportRawPreviewPdf(filteredRows, keys, headers, `${exportBasename}_raw`)}
      >
        PDF
      </button>
      <button
        type="button"
        disabled={!keys.length}
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600 disabled:opacity-40"
        onClick={() => void exportRawPreviewPptx(filteredRows, keys, headers, `${exportBasename}_raw`)}
      >
        PPTX
      </button>
      <button
        type="button"
        disabled={!keys.length}
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600 disabled:opacity-40"
        onClick={() => exportRawPreviewExcel(filteredRows, keys, headers, `${exportBasename}_raw`)}
      >
        Excel
      </button>
    </div>
  );

  const chartPanel = chartOpen ? (
    <div className="mb-6 rounded-2xl border border-slate-200/80 dark:border-slate-700/80 bg-gradient-to-b from-white/90 to-slate-50/90 dark:from-slate-900/80 dark:to-slate-950/90 p-4">
      <div className="flex flex-wrap gap-2 mb-3">
        {(["bar", "stackedBar", "line", "pie", "area"] as const).map((k) => (
          <button
            key={k}
            type="button"
            onClick={() => setChartKind(k)}
            className={`rounded-full px-3 py-1 text-[11px] font-semibold capitalize ${
              chartKind === k ? "bg-orange-500 text-white" : "bg-slate-100 dark:bg-slate-800"
            }`}
          >
            {k === "stackedBar" ? "Stack" : k}
          </button>
        ))}
      </div>
      <div
        ref={(el) => {
          (chartRef as MutableRefObject<HTMLDivElement | null>).current = el;
          if (exportChartRef) exportChartRef.current = el;
        }}
        className="h-[280px] w-full"
      >
        <ResponsiveContainer width="100%" height="100%">
          {chartKind === "pie" ? (
            <PieChart>
              <Pie
                data={pieRegion}
                dataKey="value"
                nameKey="name"
                cx="50%"
                cy="50%"
                innerRadius={44}
                outerRadius={88}
                paddingAngle={2}
              >
                {pieRegion.map((e, i) => (
                  <Cell key={i} fill={e.fill} stroke={isDark ? "#0f172a" : "#fff"} />
                ))}
              </Pie>
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
            </PieChart>
          ) : chartKind === "line" ? (
            <LineChart data={regionChartRows}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Line type="monotone" dataKey="value" stroke="#f97316" strokeWidth={2} dot name="Lignes" />
            </LineChart>
          ) : chartKind === "area" ? (
            <AreaChart data={regionChartRows}>
              <defs>
                <linearGradient id="raw-area-g" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor="#f97316" stopOpacity={0.85} />
                  <stop offset="100%" stopColor="#f97316" stopOpacity={0.08} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Area type="monotone" dataKey="value" stroke="#f97316" fill="url(#raw-area-g)" name="Lignes" />
            </AreaChart>
          ) : (
            <BarChart data={regionChartRows} margin={{ bottom: 8 }}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: palette.muted }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Bar
                dataKey="value"
                radius={[8, 8, 0, 0]}
                stackId={chartKind === "stackedBar" ? "s" : undefined}
                maxBarSize={40}
              >
                {regionChartRows.map((e, i) => (
                  <Cell key={i} fill={e.fill} />
                ))}
              </Bar>
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  ) : null;

  return (
    <GlassShell
      title="Liste des appels et des tickets"
      subtitle="Consultez les appels et les tickets correspondant aux critères sélectionnés."
      toolbar={toolbar}
    >
      {chartPanel}
      <div className="mb-3 flex flex-wrap gap-2 items-center">
        <select
          className="rounded-xl border px-3 py-1.5 text-xs bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
          value={String(rowsLimit)}
          onChange={(e) => setRowsLimit(e.target.value === "ALL" ? "ALL" : Number(e.target.value))}
        >
          {[100, 200, 500, 1000].map((n) => (
            <option key={n} value={n}>
              {n}
            </option>
          ))}
          <option value="ALL">ALL</option>
        </select>
        <input
          type="date"
          value={dateMin}
          onChange={(e) => setDateMin(e.target.value)}
          className="rounded-xl border px-3 py-1.5 text-xs bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
        />
        <input
          type="date"
          value={dateMax}
          onChange={(e) => setDateMax(e.target.value)}
          className="rounded-xl border px-3 py-1.5 text-xs bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
        />
        <input
          type="search"
          placeholder="Recherche globale..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          className="min-w-[220px] rounded-xl border px-3 py-1.5 text-xs bg-white/80 dark:bg-slate-900/70 dark:border-slate-600"
        />
        <span className="text-xs text-slate-500">{filteredRows.length} lignes filtrées</span>
      </div>
      {!visibleDefs.length ? (
        <p className="text-sm text-slate-500">Activez au moins une colonne (bouton Colonnes).</p>
      ) : (
        <div
          ref={(el) => {
            (tableRef as MutableRefObject<HTMLDivElement | null>).current = el;
            if (exportTableRef) exportTableRef.current = el;
          }}
          className="overflow-auto max-h-[420px] rounded-2xl border border-slate-200 dark:border-slate-700 shadow-inner bg-white/50 dark:bg-slate-950/40"
        >
          <table className="min-w-[760px] w-full text-[11px]">
            <thead className="bg-slate-900 text-white sticky top-0 z-10">
              <tr>
                {visibleDefs.map((c) => (
                  <th key={c.key} className="px-3 py-2 text-left whitespace-nowrap">
                    {c.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((r, i) => (
                <tr key={`${r.rawIndex}-${i}`}>
                  {visibleDefs.map((col) => (
                    <td
                      key={col.key}
                      className={`px-2 py-1 border-t border-slate-100 dark:border-slate-800 ${r.valid ? "" : "bg-rose-50/80 dark:bg-rose-950/25"} ${col.tdClass ?? ""}`}
                    >
                      {col.cell(r)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </GlassShell>
  );
}
