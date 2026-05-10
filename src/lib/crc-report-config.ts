/** Single source toggles dashboard visibility + inclusion in structured exports */

export type CrcChartKey =
  | "geoBars"
  | "geoDonut"
  | "statusPie"
  | "dailyArea"
  | "monthlyBars"
  | "trendLine"
  | "teleopBars"
  | "regionCards";

export type CrcTableKey = "pivotResult" | "pivotMetier" | "pivotNature" | "teleOpStats" | "rawPreview";

export type CrcKpiKey =
  | "totalVolume"
  | "abandons"
  | "informes"
  | "tickets"
  | "teleopsDistinct"
  | "pctInformes"
  | "pctTickets"
  | "coverage";

export type RawColumnKey =
  | "date"
  | "resultatRaw"
  | "teleop"
  | "metier"
  | "nature"
  | "regionCanon"
  | "regionsource"
  | "valid"
  | "phone";

export type ExcelSheetKey =
  | "readme"
  | "kpis"
  | "flat"
  | "pivotResult"
  | "pivotMetier"
  | "pivotNature"
  | "operators"
  | "daily"
  | "monthly";

export type PdfBundleKey =
  | "kpisCards"
  | "summaryCharts"
  | "pivotResultTable"
  | "trendsCharts"
  | "metierNatureTables"
  | "teleopPage";

export type PptxSlideKey =
  | "cover"
  | "kpi"
  | "results"
  | "metier"
  | "operators"
  | "definitions";

export interface CrcReportConfig {
  version: 1;
  /** Custom report title applied to dashboard header subtitle + exports */
  reportTitle: string;
  reportSubtitle: string;
  logoDataUrl: string | null;
  charts: Record<CrcChartKey, boolean>;
  tables: Record<CrcTableKey, boolean>;
  kpis: Record<CrcKpiKey, boolean>;
  rawColumns: Record<RawColumnKey, boolean>;
  exportExcelSheets: Record<ExcelSheetKey, boolean>;
  exportPdf: Record<PdfBundleKey, boolean>;
  exportPptx: Record<PptxSlideKey, boolean>;
}

const STORAGE_KEY = "crc-report-config-v1";

const allTrueKeys = <K extends string>(
  keys: readonly K[],
): Record<K, boolean> => Object.fromEntries(keys.map((k) => [k, true])) as Record<K, boolean>;

export const CRC_CHART_KEYS = [
  "geoBars",
  "geoDonut",
  "statusPie",
  "dailyArea",
  "monthlyBars",
  "trendLine",
  "teleopBars",
  "regionCards",
] as const;

export const CRC_TABLE_KEYS = [
  "pivotResult",
  "pivotMetier",
  "pivotNature",
  "teleOpStats",
  "rawPreview",
] as const;

export const CRC_KPI_KEYS = [
  "totalVolume",
  "abandons",
  "informes",
  "tickets",
  "teleopsDistinct",
  "pctInformes",
  "pctTickets",
  "coverage",
] as const;

export const RAW_COLUMN_KEYS = [
  "date",
  "resultatRaw",
  "teleop",
  "metier",
  "nature",
  "regionCanon",
  "regionsource",
  "valid",
  "phone",
] as const;

function defaultReportConfig(): CrcReportConfig {
  return {
    version: 1,
    reportTitle: "",
    reportSubtitle: "",
    logoDataUrl: null,
    charts: allTrueKeys(CRC_CHART_KEYS),
    tables: allTrueKeys(CRC_TABLE_KEYS),
    kpis: allTrueKeys(CRC_KPI_KEYS),
    rawColumns: allTrueKeys(RAW_COLUMN_KEYS),
    exportExcelSheets: allTrueKeys([
      "readme",
      "kpis",
      "flat",
      "pivotResult",
      "pivotMetier",
      "pivotNature",
      "operators",
      "daily",
      "monthly",
    ]),
    exportPdf: allTrueKeys([
      "kpisCards",
      "summaryCharts",
      "pivotResultTable",
      "trendsCharts",
      "metierNatureTables",
      "teleopPage",
    ]),
    exportPptx: allTrueKeys(["cover", "kpi", "results", "metier", "operators", "definitions"]),
  };
}

function parseStored(raw: string | null): CrcReportConfig {
  const fallback = defaultReportConfig();
  if (!raw) return fallback;
  try {
    const o = JSON.parse(raw) as Partial<CrcReportConfig>;
    if (o.version !== 1) return fallback;
    return {
      ...fallback,
      ...o,
      charts: { ...fallback.charts, ...o.charts },
      tables: { ...fallback.tables, ...o.tables },
      kpis: { ...fallback.kpis, ...o.kpis },
      rawColumns: { ...fallback.rawColumns, ...o.rawColumns },
      exportExcelSheets: { ...fallback.exportExcelSheets, ...o.exportExcelSheets },
      exportPdf: { ...fallback.exportPdf, ...o.exportPdf },
      exportPptx: { ...fallback.exportPptx, ...o.exportPptx },
    };
  } catch {
    return fallback;
  }
}

export function loadCrcReportConfig(): CrcReportConfig {
  if (typeof window === "undefined") return defaultReportConfig();
  return parseStored(localStorage.getItem(STORAGE_KEY));
}

export function saveCrcReportConfig(cfg: CrcReportConfig): void {
  if (typeof window === "undefined") return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    window.dispatchEvent(new Event("crc-report-config"));
  } catch {
    /** quota / privacy mode */
  }
}

export function resetCrcReportConfig(): CrcReportConfig {
  const fresh = defaultReportConfig();
  saveCrcReportConfig(fresh);
  return fresh;
}

export const defaultCrcReportConfig = defaultReportConfig;

export function mergeExportPdf(plan?: Partial<CrcReportConfig["exportPdf"]>) {
  return { ...defaultReportConfig().exportPdf, ...plan };
}

export function mergeExportExcelSheets(plan?: Partial<CrcReportConfig["exportExcelSheets"]>) {
  return { ...defaultReportConfig().exportExcelSheets, ...plan };
}

export function mergeExportPptx(plan?: Partial<CrcReportConfig["exportPptx"]>) {
  return { ...defaultReportConfig().exportPptx, ...plan };
}
