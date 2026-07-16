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
import { getResultColor } from "@/lib/constants/chart-colors";
import type { RegionPivotId } from "@/lib/crc-export-helpers";
import { activeRegionShorts, numericValue } from "@/lib/crc-export-helpers";
import {
  captureElementPng,
  exportMatrixExcel,
  exportMatrixPdf,
  exportMatrixPptx,
} from "@/lib/crc-export-engine";
import type { PivotRegionRow } from "@/lib/crc-export-helpers";

export type PivotChartKind = "bar" | "stackedBar" | "line" | "pie" | "area";

function GlassCard({
  title,
  subtitle,
  children,
  toolbar,
}: {
  title: string;
  subtitle?: string;
  children: ReactNode;
  toolbar?: ReactNode;
}) {
  return (
    <section className="glass-panel p-5 sm:p-6 hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-slate-900/10 dark:hover:shadow-slate-950/40 motion-safe:animate-fade-in">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between mb-4">
        <div>
          <h3 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50">{title}</h3>
          {subtitle ? <p className="text-sm text-slate-600 dark:text-slate-400 mt-0.5">{subtitle}</p> : null}
        </div>
      </div>
      {toolbar}
      {children}
    </section>
  );
}

export function CrcRegionPivotWidget(props: {
  widgetId: RegionPivotId;
  title: string;
  subtitle?: string;
  labelHeader: string;
  /** First column accessor key in row object */
  rowLabelKey: "name" | "métier" | "nature";
  rows: PivotRegionRow[];
  regionVisibility: Record<string, boolean>;
  onToggleRegion: (regionShort: string) => void;
  isDark: boolean;
  palette: { fg: string; muted: string; grid: string; tooltipBg: string };
  exportBasename: string;
  /** Pour inclusion PDF/PPTX (capture du tableau / graphique ouvert) */
  exportTableRef?: RefObject<HTMLDivElement | null>;
  exportChartRef?: RefObject<HTMLDivElement | null>;
}) {
  const {
    title,
    subtitle,
    labelHeader,
    rowLabelKey,
    rows,
    regionVisibility,
    onToggleRegion,
    isDark,
    palette,
    exportBasename,
    exportTableRef,
    exportChartRef,
  } = props;
  const [colsOpen, setColsOpen] = useState(false);
  const [chartOpen, setChartOpen] = useState(false);
  const [chartKind, setChartKind] = useState<PivotChartKind>("bar");
  const chartRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  const shorts = useMemo(() => {
    const allShorts = activeRegionShorts(regionVisibility);
    return allShorts.filter((s) => rows.some((row) => numericValue(row, s) > 0));
  }, [regionVisibility, rows]);

  const chartRows = useMemo(() => {
    return rows.map((row) => {
      const label = String(row[rowLabelKey] ?? row.name ?? "");
      const o: Record<string, string | number> = { label };
      for (const s of shorts) {
        o[s] = numericValue(row, s);
      }
      return o;
    });
  }, [rows, rowLabelKey, shorts]);

  const pieAgg = useMemo(() => {
    const m = new Map<string, number>();
    shorts.forEach((s) => m.set(s, 0));
    rows.forEach((row) => {
      shorts.forEach((s) => {
        m.set(s, (m.get(s) ?? 0) + numericValue(row, s));
      });
    });
    return shorts.map((s) => {
      const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === s) ?? REGION_ORDER[0];
      return {
        name: s,
        value: m.get(s) ?? 0,
        fill: REGION_COLORS[rg],
      };
    });
  }, [rows, shorts]);

  const regionColor = (s: string) =>
    REGION_COLORS[REGION_ORDER.find((rg) => REGION_SHORT[rg] === s) ?? REGION_ORDER[0]];

  const toolbar = (
    <div className="flex flex-wrap items-center gap-2 mb-4 pb-3 border-b border-slate-200/70 dark:border-slate-700/60">
      <div className="relative">
        <button
          type="button"
          onClick={() => setColsOpen((o) => !o)}
          className="rounded-full px-4 py-2 text-xs font-semibold bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-600 hover:brightness-95"
        >
          Colonnes
        </button>
        {colsOpen ? (
          <div className="absolute left-0 z-30 mt-2 w-56 rounded-2xl border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-900 shadow-xl p-3 space-y-2 max-h-64 overflow-y-auto">
            {REGION_ORDER.map((rg) => {
              const s = REGION_SHORT[rg];
              return (
                <label key={s} className="flex items-center gap-2 text-xs cursor-pointer">
                  <input
                    type="checkbox"
                    checked={regionVisibility[s] !== false}
                    onChange={() => onToggleRegion(s)}
                  />
                  <span style={{ color: REGION_COLORS[rg] }}>{s}</span>
                </label>
              );
            })}
          </div>
        ) : null}
      </div>
      <button
        type="button"
        onClick={() => setChartOpen((c) => !c)}
        className={`rounded-full px-4 py-2 text-xs font-semibold border ${
          chartOpen
            ? "border-orange-500 bg-orange-500/15 text-orange-900 dark:text-orange-100"
            : "border-slate-200 dark:border-slate-600 bg-white/80 dark:bg-slate-800"
        }`}
      >
        Charts
      </button>
      <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 px-1">Exporter</span>
      <span className="text-slate-300 dark:text-slate-600">|</span>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={async () => {
          const el = chartOpen && chartRef.current ? chartRef.current : tableRef.current;
          if (!el) return;
          await captureElementPng(el, `${exportBasename}_${props.widgetId}`);
        }}
      >
        PNG
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportMatrixPdf(title, labelHeader, rows, shorts, `${exportBasename}_${props.widgetId}`)}
      >
        PDF
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportMatrixPptx(title, labelHeader, rows, shorts, `${exportBasename}_${props.widgetId}`)}
      >
        PPTX
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => exportMatrixExcel(title, labelHeader, rows, shorts, `${exportBasename}_${props.widgetId}`)}
      >
        Excel
      </button>
    </div>
  );

  const chartPanel = chartOpen ? (
    <div className="mb-6 space-y-3 rounded-2xl border border-slate-200/80 dark:border-slate-700/80 bg-gradient-to-b from-white/90 to-slate-50/90 dark:from-slate-900/80 dark:to-slate-950/90 p-4">
      <div className="flex flex-wrap gap-2">
        {(["bar", "stackedBar", "line", "pie", "area"] as const).map((k) => (
          <button
            key={k}
            type="button"
            onClick={() => setChartKind(k)}
            className={`rounded-full px-3 py-1 text-[11px] font-semibold capitalize ${
              chartKind === k ? "bg-orange-500 text-white shadow" : "bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-200"
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
        className="h-[320px] w-full"
      >
        <ResponsiveContainer width="100%" height="100%">
          {chartKind === "pie" ? (
            <PieChart>
              <Pie data={pieAgg} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={50} outerRadius={100} paddingAngle={2}>
                {pieAgg.map((e, i) => (
                  <Cell key={i} fill={e.fill} stroke={isDark ? "#0f172a" : "#fff"} strokeWidth={1} />
                ))}
              </Pie>
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
            </PieChart>
          ) : chartKind === "line" ? (
            <LineChart data={chartRows}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted, fontSize: 10 }} hide={chartRows.length > 16} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
              {shorts.map((s) => (
                <Line
                  key={s}
                  type="monotone"
                  dataKey={s}
                  stroke={regionColor(s)}
                  strokeWidth={2}
                  dot={false}
                  name={s}
                />
              ))}
            </LineChart>
          ) : chartKind === "area" ? (
            <AreaChart data={chartRows}>
              <defs>
                {shorts.map((s) => (
                  <linearGradient key={s} id={`g-${props.widgetId}-${s}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={regionColor(s)} stopOpacity={0.85} />
                    <stop offset="100%" stopColor={regionColor(s)} stopOpacity={0.08} />
                  </linearGradient>
                ))}
              </defs>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted, fontSize: 10 }} hide={chartRows.length > 16} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
              {shorts.map((s) => (
                <Area
                  key={s}
                  type="monotone"
                  dataKey={s}
                  stackId="a"
                  stroke={regionColor(s)}
                  fill={`url(#g-${props.widgetId}-${s})`}
                  name={s}
                />
              ))}
            </AreaChart>
          ) : (
            <BarChart data={chartRows} margin={{ bottom: chartRows.length > 12 ? 48 : 8 }}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" vertical={false} />
              <XAxis
                dataKey="label"
                tick={{ fill: palette.muted, fontSize: 10 }}
                angle={chartRows.length > 10 ? -22 : 0}
                textAnchor={chartRows.length > 10 ? "end" : "middle"}
                height={chartRows.length > 10 ? 56 : 28}
              />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
              {shorts.map((s) => (
                <Bar
                  key={s}
                  dataKey={s}
                  fill={regionColor(s)}
                  radius={[6, 6, 0, 0]}
                  stackId={chartKind === "stackedBar" ? "st" : undefined}
                  maxBarSize={32}
                />
              ))}
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  ) : null;

  return (
    <GlassCard title={title} subtitle={subtitle} toolbar={toolbar}>
      {chartPanel}
      <div
        ref={(el) => {
          (tableRef as MutableRefObject<HTMLDivElement | null>).current = el;
          if (exportTableRef) exportTableRef.current = el;
        }}
        className="overflow-x-auto rounded-2xl border border-slate-200/70 dark:border-slate-700/80"
      >
        <table className="crc min-w-[560px] w-full text-xs sm:text-sm">
          <thead className="bg-slate-900 text-white dark:bg-slate-950 sticky top-0 z-10 shadow">
            <tr>
              <th className="px-3 py-2 text-left font-semibold">{labelHeader}</th>
              {shorts.map((s) => {
                const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === s) ?? REGION_ORDER[0];
                return (
                  <th
                    key={s}
                    className="px-3 py-2 text-center font-semibold whitespace-nowrap"
                    style={{ backgroundColor: `${REGION_COLORS[rg]}cc`, color: "#fff" }}
                  >
                    {s}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => {
              const label = String(row[rowLabelKey] ?? row.name ?? "");
              const leftColor = rowLabelKey === "name" ? getResultColor(label) : undefined;
              return (
                <tr
                  key={`${idx}-${label}`}
                  className={idx % 2 === 0 ? "bg-white/95 dark:bg-slate-950/55" : "bg-slate-50 dark:bg-slate-900/50"}
                >
                  <td
                    className="px-3 py-1.5 font-medium text-slate-900 dark:text-slate-100 border-l-[4px] pl-2.5"
                    style={{ borderLeftColor: leftColor ?? "#94a3b8" }}
                  >
                    {label}
                  </td>
                  {shorts.map((s) => (
                    <td key={s} className="px-3 py-1.5 text-center tabular-nums">
                      {numericValue(row, s)}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </GlassCard>
  );
}
