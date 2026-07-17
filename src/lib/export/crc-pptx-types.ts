import type { CrcExportColumnVisibility } from "../crc-export-helpers";
import type { CrcChartKey, CrcKpiKey, CrcReportConfig, CrcTableKey, RawColumnKey } from "../crc-report-config";

/** Live dashboard snapshot — same toggles as the dashboard. */
export type CrcPptxDashboardSnapshot = Pick<
  CrcReportConfig,
  "charts" | "tables" | "kpis" | "rawColumns" | "exportPptx"
>;

/** Extensible widget description for the PPTX engine (future widgets). */
export type CrcPptxWidgetKind = "cover" | "kpi" | "chart" | "table" | "definitions" | "composite";

export interface CrcPptxWidgetDescriptor {
  id: string;
  kind: CrcPptxWidgetKind;
  title: string;
  subtitle?: string;
  /** Keys from dashboard config that gate this block */
  chartKeys?: CrcChartKey[];
  tableKeys?: CrcTableKey[];
  kpiKeys?: CrcKpiKey[];
  rawColumnKeys?: RawColumnKey[];
  meta?: Record<string, unknown>;
}

export type CrcPptxExportOptions = {
  subtitle?: string;
  logoOverride?: string | null;
  columnVisibility?: CrcExportColumnVisibility;
  /** Required — mirrors visible dashboard widgets */
  dashboard: Pick<CrcReportConfig, "charts" | "tables" | "kpis" | "rawColumns" | "exportPptx">;
  pptx?: Partial<CrcReportConfig["exportPptx"]>;
};
