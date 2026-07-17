/**
 * Central registry of exportable CRC dashboard widgets.
 * PPTX builder and per-widget exporters derive payloads from here to stay in sync.
 */

import { regionSlice } from "../crc-analytics";
import { getResultColor } from "../constants/chart-colors";
import type { CanonicalRegion } from "../crc-constants";
import { REGION_ORDER, REGION_SHORT } from "../crc-constants";
import type { CrcPptxDashboardSnapshot } from "./crc-pptx-types";
import type { CrcRow } from "../crc-types";

/** Région — histogramme résultats (carte dashboard). */
export interface CrcRegionCardWidgetSpec {
  widgetId: string;
  widgetTitle: string;
  widgetType: "chart";
  chartType: "bar";
  chartLabels: string[];
  chartDataset: number[];
  /** CSS hex for UI / reference */
  chartColors: string[];
  region: CanonicalRegion;
  visible: boolean;
  exportEnabled: boolean;
}

export interface CrcPptxWidgetRegistry {
  regionCards: CrcRegionCardWidgetSpec[];
}

export function buildRegionCardWidgetSpecs(
  rows: CrcRow[],
  dashboard: Pick<CrcPptxDashboardSnapshot, "charts">,
): CrcRegionCardWidgetSpec[] {
  if (!dashboard.charts.regionCards || rows.length === 0) return [];

  return REGION_ORDER.map((rg) => {
    const a = regionSlice(rows, rg);
    const pairs = a.résultat.slice(0, 8);
    const chartLabels = pairs.map(([n]) => n);
    const chartDataset = pairs.map(([, c]) => c);
    const chartColors = chartLabels.map((lbl) => getResultColor(lbl));
    return {
      widgetId: `region-${rg}-resultats`,
      widgetTitle: `${REGION_SHORT[rg]} — Résultats`,
      widgetType: "chart" as const,
      chartType: "bar" as const,
      chartLabels,
      chartDataset,
      chartColors,
      region: rg,
      visible: true,
      exportEnabled: true,
    };
  });
}

/** Snapshot consumed by `pptx-builder` (extend with tables, KPIs, etc.). */
export function buildCrcPptxWidgetRegistry(rows: CrcRow[], dash: CrcPptxDashboardSnapshot): CrcPptxWidgetRegistry {
  return {
    regionCards: buildRegionCardWidgetSpecs(rows, dash),
  };
}
