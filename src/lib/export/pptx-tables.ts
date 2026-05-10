import type PptxGenJS from "pptxgenjs";

import {
  AGGREGATE_VOLUME_BAR,
  RESULT_COLORS,
  getResultColor,
  hexForPptx,
} from "../constants/chart-colors";
import { REGION_COLORS, REGION_ORDER, REGION_SHORT } from "../crc-constants";
import type { PivotRegionRow } from "../crc-export-helpers";
import { formatRawCellForExport, numericValue } from "../crc-export-helpers";
import type { RawColumnKey } from "../crc-report-config";
import type { CrcRow } from "../crc-types";

import { PPTX_CRC } from "./pptx-theme";

function matrixRowLabel(row: PivotRegionRow): string {
  const rec = row as Record<string, unknown>;
  if (typeof rec.name === "string" && rec.name.length) return rec.name;
  if (typeof rec.métier === "string") return String(rec.métier);
  if (typeof rec.nature === "string") return String(rec.nature);
  return "";
}

export function pivotTableRows(
  rows: PivotRegionRow[],
  regionShorts: string[],
  labelHeader: string,
): (string | { text: string; options?: { color?: string } })[][] {
  const head: { text: string; options: { bold: boolean; fill: { color: string }; color: string } }[] = [
    {
      text: labelHeader,
      options: { bold: true, fill: { color: PPTX_CRC.navy }, color: PPTX_CRC.white },
    },
    ...regionShorts.map((s) => {
      const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === s);
      return {
        text: s,
        options: {
          bold: true,
          fill: { color: rg ? hexForPptx(REGION_COLORS[rg]) : "64748b" },
          color: PPTX_CRC.white,
        },
      };
    }),
    {
      text: "Total",
      options: { bold: true, fill: { color: PPTX_CRC.navy }, color: PPTX_CRC.white },
    },
  ];

  const body = rows.map((row) => {
    const label = matrixRowLabel(row);
    const nums = regionShorts.map((s) => numericValue(row, s));
    const tot = nums.reduce((a, b) => a + b, 0);
    const firstCell =
      labelHeader === "Résultat"
        ? { text: label, options: { color: hexForPptx(getResultColor(label)) } }
        : { text: label };
    return [firstCell, ...nums.map((n) => String(n)), String(tot)];
  });

  return [head, ...body] as (string | { text: string; options?: { color?: string } })[][];
}

export function addPivotTable(
  slide: PptxGenJS.Slide,
  rows: PivotRegionRow[],
  regionShorts: string[],
  labelHeader: string,
  layout: { x: number; y: number; w: number; fontSize?: number },
) {
  const data = pivotTableRows(rows, regionShorts, labelHeader);
  slide.addTable(data as never, {
    x: layout.x,
    y: layout.y,
    w: layout.w,
    fontSize: layout.fontSize ?? 8,
    border: { pt: 0.5, color: PPTX_CRC.cardBorder },
  });
}

export function teleopTableRows(
  ops: {
    name: string;
    volume: number;
    abandons: number;
    appelsDécrochésInterrompus: number;
    informés: number;
    tickets: number;
  }[],
  metricKeys: string[],
) {
  const short: Record<string, string> = {
    volume: "Volume",
    abandons: "Abandons",
    appelsDécrochésInterrompus: "Décrochés",
    informés: "Informés",
    tickets: "Tickets",
  };
  const teleColor: Record<string, string> = {
    volume: AGGREGATE_VOLUME_BAR,
    abandons: RESULT_COLORS["Appels abandonnés"],
    appelsDécrochésInterrompus: RESULT_COLORS["Appels décrochés interrompus"],
    informés: RESULT_COLORS["Clients informés"],
    tickets: RESULT_COLORS["Tickets transmis"],
  };

  const head: PptxGenJS.TableCell[] = [
    { text: "Téléopérateur", options: { bold: true, fill: { color: PPTX_CRC.navy }, color: PPTX_CRC.white } },
    ...metricKeys.map((key) => ({
      text: short[key] ?? key,
      options: { bold: true, fill: { color: PPTX_CRC.navy }, color: PPTX_CRC.white },
    })),
  ];

  const body = ops.map((o) => [
    { text: o.name },
    ...metricKeys.map((key) => ({
      text: String(numericValue(o, key)),
      options: { color: hexForPptx(teleColor[key] ?? "#64748b") },
    })),
  ]);

  return [head, ...body];
}

export function addRawPreviewTable(
  slide: PptxGenJS.Slide,
  rows: CrcRow[],
  keys: RawColumnKey[],
  headers: string[],
  layout: { x: number; y: number; w: number; maxRows?: number },
) {
  const max = layout.maxRows ?? 24;
  const head = headers.map((h) => ({
    text: h,
    options: { bold: true, fill: { color: PPTX_CRC.navy }, color: PPTX_CRC.white },
  }));
  const body = rows.slice(0, max).map((r) =>
    keys.map((k) => ({
      text: formatRawCellForExport(r, k),
    })),
  );
  slide.addTable([head, ...body] as never, {
    x: layout.x,
    y: layout.y,
    w: layout.w,
    fontSize: 7,
    border: { pt: 0.5, color: PPTX_CRC.cardBorder },
  });
}
