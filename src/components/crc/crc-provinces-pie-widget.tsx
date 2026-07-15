"use client";

import { useMemo } from "react";
import {
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
} from "recharts";

import { provincesParRégion } from "@/lib/crc-analytics";
import {
  REGION_COLORS,
  REGION_ORDER,
  REGION_SHORT,
  type CanonicalRegion,
} from "@/lib/crc-constants";
import type { CrcRow } from "@/lib/crc-types";

interface ProvincesPieWidgetProps {
  rows: CrcRow[];
  palette: {
    fg: string;
    muted: string;
    grid: string;
    tooltipBg: string;
    series: string[];
  };
}

function GlassCard({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <section className="glass-panel p-5 sm:p-6 hover:-translate-y-0.5 hover:shadow-2xl hover:shadow-slate-900/10 dark:hover:shadow-slate-950/40 motion-safe:animate-fade-in">
      <h3 className="text-lg font-semibold tracking-tight text-slate-900 dark:text-slate-50 mb-4">
        {title}
      </h3>
      {children}
    </section>
  );
}

export function CrcProvincesPieWidget(props: ProvincesPieWidgetProps) {
  const { rows, palette } = props;

  const provincesData = useMemo(() => provincesParRégion(rows), [rows]);

  const chartTooltip = (
    <Tooltip
      contentStyle={{
        backgroundColor: palette.tooltipBg,
        borderRadius: 12,
        border: `1px solid ${palette.grid}`,
      }}
      labelStyle={{ color: palette.fg }}
      formatter={(value: number, name: string) => [`${Number(value)?.toLocaleString("fr-FR") ?? 0}`, name]}
    />
  );

  return (
    <section className="grid grid-cols-1 lg:grid-cols-3 gap-4">
      {REGION_ORDER.slice(0, 3).map((region) => {
        const provinces = provincesData[region] || [];
        const pieData = provinces.map(([name, value]) => ({
          name: name || "Non renseigné",
          value,
        }));

        return (
          <GlassCard key={region} title={`Provinces — ${REGION_SHORT[region]}`}>
            {pieData.length > 0 ? (
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={pieData}
                      dataKey="value"
                      cx="50%"
                      cy="50%"
                      outerRadius={80}
                      stroke="none"
                    >
                      {pieData.map((entry, idx) => (
                        <Cell
                          key={`${region}-${idx}`}
                          fill={palette.series[idx % palette.series.length]}
                          opacity={0.9}
                        />
                      ))}
                    </Pie>
                    <Legend
                      formatter={(value) => (
                        <span style={{ color: palette.fg, fontSize: 11 }}>
                          {value}
                        </span>
                      )}
                      wrapperStyle={{ paddingTop: 8 }}
                    />
                    {chartTooltip}
                  </PieChart>
                </ResponsiveContainer>
              </div>
            ) : (
              <div className="h-72 flex items-center justify-center text-slate-500 dark:text-slate-400">
                Aucune donnée de province pour cette région
              </div>
            )}
          </GlassCard>
        );
      })}
    </section>
  );
}