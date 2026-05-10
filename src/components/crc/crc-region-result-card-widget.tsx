"use client";

import { useMemo, useRef, type MutableRefObject, type ReactNode, type RefObject } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";

import { getResultColor } from "@/lib/constants/chart-colors";
import { regionSlice } from "@/lib/crc-analytics";
import type { CanonicalRegion } from "@/lib/crc-constants";
import { REGION_SHORT } from "@/lib/crc-constants";
import {
  captureElementPng,
  exportRegionResultBarExcel,
  exportRegionResultBarPdf,
  exportRegionResultBarPptx,
  getRegionResultChartRows,
} from "@/lib/crc-export-engine";
import type { CrcRow } from "@/lib/crc-types";

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

export function CrcRegionResultCardWidget(props: {
  region: CanonicalRegion;
  rows: CrcRow[];
  palette: { fg: string; muted: string; grid: string; tooltipBg: string };
  exportBasename: string;
  exportChartRef?: RefObject<HTMLDivElement | null>;
}) {
  const { region, rows, palette, exportBasename, exportChartRef } = props;
  const chartRef = useRef<HTMLDivElement>(null);

  const analytics = useMemo(() => regionSlice(rows, region), [rows, region]);
  const chartData = useMemo(
    () => analytics.résultat.slice(0, 8).map(([name, cnt]) => ({ name, cnt })),
    [analytics.résultat],
  );
  const chartRows = useMemo(() => getRegionResultChartRows(rows, region), [rows, region]);

  const widgetId = `region-${region}-resultats`;
  const widgetTitle = `${REGION_SHORT[region]} — Résultats`;
  const exportSlug = `${exportBasename}_${widgetId}`;

  const chartTooltip = (
    <Tooltip
      contentStyle={{
        backgroundColor: palette.tooltipBg,
        borderRadius: 12,
        border: `1px solid ${palette.grid}`,
      }}
      labelStyle={{ color: palette.fg }}
      formatter={(value: number) => [`${Number(value)?.toLocaleString("fr-FR") ?? 0}`, "Interactions"]}
    />
  );

  const toolbar = (
    <div className="flex flex-wrap items-center gap-2 mb-4 pb-3 border-b border-slate-200/70 dark:border-slate-700/60">
      <span className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 px-1">Exporter</span>
      <span className="text-slate-300 dark:text-slate-600">|</span>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={async () => {
          const el = chartRef.current;
          if (!el) return;
          await captureElementPng(el, exportSlug);
        }}
      >
        PNG
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportRegionResultBarPdf(widgetTitle, chartRows, exportSlug)}
      >
        PDF
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => void exportRegionResultBarPptx(widgetTitle, chartRows, exportSlug)}
      >
        PPTX
      </button>
      <button
        type="button"
        className="rounded-full px-3 py-2 text-[11px] font-semibold border border-slate-200 dark:border-slate-600"
        onClick={() => exportRegionResultBarExcel(widgetTitle, chartRows, exportSlug)}
      >
        Excel
      </button>
    </div>
  );

  return (
    <GlassShell title={REGION_SHORT[region]} subtitle={`${analytics.total} interactions filtrées`} toolbar={toolbar}>
      <div
        ref={(el) => {
          (chartRef as MutableRefObject<HTMLDivElement | null>).current = el;
          if (exportChartRef) exportChartRef.current = el;
        }}
        className="h-52 mb-4"
      >
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={chartData}>
            <CartesianGrid stroke={palette.grid} vertical={false} />
            <XAxis dataKey="name" hide tick={false} />
            <YAxis tick={{ fill: palette.muted, fontSize: 10 }} />
            {chartTooltip}
            <Bar dataKey="cnt" radius={[6, 6, 4, 4]}>
              {chartData.map((d) => (
                <Cell key={d.name} fill={getResultColor(d.name)} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="text-[11px] space-y-1 text-slate-600 dark:text-slate-400">
        <p className="font-semibold text-slate-800 dark:text-slate-100">Nature top 5</p>
        <ul className="list-disc ml-5">
          {analytics.nature.slice(0, 5).map(([k, v]) => (
            <li key={k}>
              <span className="font-medium">{k}</span> ({v})
            </li>
          ))}
        </ul>
      </div>
    </GlassShell>
  );
}
