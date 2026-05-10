import type { CrcRow } from "./crc-types";
import { defaultCrcReportConfig } from "./crc-report-config";
import { exportCrcExcel } from "./export-excel";
import { exportCrcPdf } from "./export-pdf";
import { exportCrcPowerPoint } from "./export-pptx";

export async function exportWidgetAsPdf(rows: CrcRow[], title: string) {
  await exportCrcPdf(rows, title, `${title.replace(/\s+/g, "_")}_widget`);
}

export async function exportWidgetAsPptx(rows: CrcRow[], title: string) {
  const d = defaultCrcReportConfig();
  await exportCrcPowerPoint(rows, title, `${title.replace(/\s+/g, "_")}_widget`, {
    dashboard: {
      charts: d.charts,
      tables: d.tables,
      kpis: d.kpis,
      rawColumns: d.rawColumns,
      exportPptx: d.exportPptx,
    },
  });
}

export function exportWidgetAsExcel(rows: CrcRow[], title: string) {
  exportCrcExcel(rows, `${title.replace(/\s+/g, "_")}_widget`);
}

export function exportWidgetAsPngFallback() {
  /** Placeholder until per-widget SVG/canvas capture is plugged in. */
  throw new Error("PNG export hook not yet bound to widget renderer.");
}
