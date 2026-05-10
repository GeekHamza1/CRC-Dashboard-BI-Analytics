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

import { AGGREGATE_VOLUME_BAR, RESULT_COLORS } from "@/lib/constants/chart-colors";
import type { OperatorRankRow, TeleOpMetricKey } from "@/lib/crc-export-helpers";
import { activeTeleOpKeys } from "@/lib/crc-export-helpers";
import { captureElementPng, exportTeleOpExcel, exportTeleOpPdf, exportTeleOpPptx } from "@/lib/crc-export-engine";

import type { PivotChartKind } from "./crc-region-pivot-widget";

const METRIC_LABEL: Record<TeleOpMetricKey, string> = {
  volume: "Volume",
  abandons: "Appels abandonnés",
  appelsDécrochésInterrompus: "Appels décrochés interrompus",
  informés: "Clients informés",
  tickets: "Tickets transmis",
};

const METRIC_COLOR: Record<TeleOpMetricKey, string> = {
  volume: AGGREGATE_VOLUME_BAR,
  abandons: RESULT_COLORS["Appels abandonnés"],
  appelsDécrochésInterrompus: RESULT_COLORS["Appels décrochés interrompus"],
  informés: RESULT_COLORS["Clients informés"],
  tickets: RESULT_COLORS["Tickets transmis"],
};

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

export function CrcTeleopStatsWidget(props: {
  rows: OperatorRankRow[];
  metricVisibility: Record<string, boolean>;
  onToggleMetric: (key: TeleOpMetricKey) => void;
  isDark: boolean;
  palette: { muted: string; grid: string; tooltipBg: string };
  exportBasename: string;
  exportTableRef?: RefObject<HTMLDivElement | null>;
  exportChartRef?: RefObject<HTMLDivElement | null>;
}) {
  const { rows, metricVisibility, onToggleMetric, isDark, palette, exportBasename, exportTableRef, exportChartRef } =
    props;
  const [colsOpen, setColsOpen] = useState(false);
  const [chartOpen, setChartOpen] = useState(false);
  const [chartKind, setChartKind] = useState<PivotChartKind>("bar");
  const chartRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  const keys = useMemo(() => activeTeleOpKeys(metricVisibility), [metricVisibility]);

  const chartData = useMemo(() => {
    return rows.slice(0, 14).map((o) => {
      const label = o.name.length > 16 ? `${o.name.slice(0, 15)}…` : o.name;
      const d: Record<string, string | number> = { label };
      keys.forEach((k) => {
        d[k] = (o as Record<string, number>)[k] ?? 0;
      });
      return d;
    });
  }, [rows, keys]);

  const pieTotals = useMemo(() => {
    const acc: Record<string, number> = {};
    keys.forEach((k) => {
      acc[k] = rows.reduce((s, o) => s + Number((o as Record<string, number>)[k] ?? 0), 0);
    });
    return keys.map((k) => ({
      name: METRIC_LABEL[k],
      value: acc[k] ?? 0,
      fill: METRIC_COLOR[k],
    }));
  }, [rows, keys]);

  const metricKeysList = (["volume", "abandons", "appelsDécrochésInterrompus", "informés", "tickets"] as const);

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
          <div className="absolute left-0 z-30 mt-2 w-64 rounded-2xl border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-900 shadow-xl p-3 space-y-2">
            {metricKeysList.map((k) => (
              <label key={k} className="flex items-center gap-2 text-xs cursor-pointer">
                <input
                  type="checkbox"
                  checked={metricVisibility[k] !== false}
                  onChange={() => onToggleMetric(k)}
                />
                <span style={{ color: METRIC_COLOR[k] }}>{METRIC_LABEL[k]}</span>
              </label>
            ))}
          </div>
        ) : null}
      </div>
      <button
        type="button"
        onClick={() => setChartOpen((c) => !c)}
        className={`rounded-full px-4 py-2 text-xs font-semibold border ${
          chartOpen ? "border-orange-500 bg-orange-500/15" : "border-slate-200 dark:border-slate-600"
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
          await captureElementPng(el, `${exportBasename}_teleop`);
        }}
      >
        PNG
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportTeleOpPdf(rows, keys as string[], `${exportBasename}_teleop`)}
      >
        PDF
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportTeleOpPptx(rows, keys as string[], `${exportBasename}_teleop`)}
      >
        PPTX
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => exportTeleOpExcel(rows, keys as string[], `${exportBasename}_teleop`)}
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
        className="h-[300px] w-full"
      >
        <ResponsiveContainer width="100%" height="100%">
          {chartKind === "pie" ? (
            <PieChart>
              <Pie data={pieTotals} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={48} outerRadius={96}>
                {pieTotals.map((e, i) => (
                  <Cell key={i} fill={e.fill} stroke={isDark ? "#0f172a" : "#fff"} />
                ))}
              </Pie>
              <Tooltip contentStyle={{ borderRadius: 12, backgroundColor: palette.tooltipBg }} />
              <Legend />
            </PieChart>
          ) : chartKind === "line" ? (
            <LineChart data={chartData}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted, fontSize: 10 }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip />
              <Legend />
              {keys.map((k) => (
                <Line key={k} type="monotone" dataKey={k} stroke={METRIC_COLOR[k]} strokeWidth={2} dot={false} name={METRIC_LABEL[k]} />
              ))}
            </LineChart>
          ) : chartKind === "area" ? (
            <AreaChart data={chartData}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={{ fill: palette.muted, fontSize: 10 }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip />
              <Legend />
              {keys.map((k) => (
                <Area
                  key={k}
                  type="monotone"
                  dataKey={k}
                  stackId="st"
                  stroke={METRIC_COLOR[k]}
                  fill={`${METRIC_COLOR[k]}55`}
                  name={METRIC_LABEL[k]}
                />
              ))}
            </AreaChart>
          ) : (
            <BarChart data={chartData}>
              <CartesianGrid stroke={palette.grid} strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: palette.muted, fontSize: 10 }} />
              <YAxis tick={{ fill: palette.muted }} />
              <Tooltip />
              <Legend />
              {keys.map((k) => (
                <Bar
                  key={k}
                  dataKey={k}
                  fill={METRIC_COLOR[k]}
                  radius={[4, 4, 0, 0]}
                  stackId={chartKind === "stackedBar" ? "b" : undefined}
                  maxBarSize={22}
                  name={METRIC_LABEL[k]}
                />
              ))}
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  ) : null;

  return (
    <GlassShell
      title="Statistiques téléopérateurs"
      subtitle="Colonnes issues du champ Résultat brut Axilus — contrôle colonnes & exports ci-dessous."
      toolbar={toolbar}
    >
      {chartPanel}
      <div
        ref={(el) => {
          (tableRef as MutableRefObject<HTMLDivElement | null>).current = el;
          if (exportTableRef) exportTableRef.current = el;
        }}
        className="overflow-x-auto rounded-2xl border border-slate-200/70 dark:border-slate-700/80"
      >
        <table className="crc min-w-[880px] w-full text-xs sm:text-sm">
          <thead className="bg-slate-900 text-white dark:bg-slate-950 sticky top-0 z-10 shadow">
            <tr>
              <th className="px-3 py-2 text-left font-semibold">Téléopérateur</th>
              {keys.map((k) => (
                <th key={k} className="px-3 py-2 text-center font-semibold whitespace-nowrap" style={{ color: "#fff" }}>
                  {METRIC_LABEL[k]}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((o, idx) => (
              <tr
                key={o.name}
                className={idx % 2 === 0 ? "bg-white/95 dark:bg-slate-950/55" : "bg-slate-50 dark:bg-slate-900/50"}
              >
                <td className="px-3 py-1.5 font-medium">{o.name}</td>
                {keys.map((k) => (
                  <td
                    key={k}
                    className="px-3 py-1.5 text-center tabular-nums font-semibold"
                    style={{ color: METRIC_COLOR[k] }}
                  >
                    {(o as Record<string, number>)[k] ?? 0}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </GlassShell>
  );
}
