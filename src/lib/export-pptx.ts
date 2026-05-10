"use client";

/**
 * CRC PowerPoint export — native PptxGenJS objects only (editable tables, charts, shapes).
 * Implementation lives in `src/lib/export/` (pptx-builder, theme, layout, KPIs, tables, charts).
 */

import type { CrcRow } from "./crc-types";
import type { CrcPptxExportOptions } from "./export/crc-pptx-types";
import { buildAndSaveCrcPptx } from "./export/pptx-builder";

export type { CrcPptxDashboardSnapshot, CrcPptxExportOptions, CrcPptxWidgetDescriptor } from "./export/crc-pptx-types";

export async function exportCrcPowerPoint(
  rows: CrcRow[],
  title = "Pilotage opérationnel CRC — Axilus",
  basename = "crc_axilus_management",
  opts?: CrcPptxExportOptions,
) {
  if (!opts?.dashboard) {
    throw new Error("exportCrcPowerPoint: `dashboard` snapshot is required (cockpit visibility).");
  }
  await buildAndSaveCrcPptx(rows, title, basename, opts);
}
