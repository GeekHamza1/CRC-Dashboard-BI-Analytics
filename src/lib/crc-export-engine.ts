"use client";

import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import PptxGenJS from "pptxgenjs";
import * as XLSX from "xlsx";

import {
  AGGREGATE_VOLUME_BAR,
  RESULT_COLORS,
  getResultColor,
  hexForPptx,
  hexToRgbTuple,
} from "./constants/chart-colors";
import { resolveReportLogo, reportTimestamp } from "./export-branding";
import type { RawColumnKey } from "./crc-report-config";
import type { CrcRow } from "./crc-types";
import { regionSlice } from "./crc-analytics";
import type { CanonicalRegion } from "./crc-constants";
import { formatRawCellForExport, numericValue, type OperatorRankRow, type PivotRegionRow } from "./crc-export-helpers";
import { addResultDistributionBarChart } from "./export/pptx-charts";

function sanitizeBase(s: string) {
  return s.replace(/[^\wÀ-ÖØ-öø-ÿ\-]+/gi, "_").slice(0, 120);
}

async function nodeToPngDataUrl(node: HTMLElement) {
  const { toPng } = await import("html-to-image");
  return await toPng(node, {
    cacheBust: true,
    backgroundColor:
      typeof document !== "undefined" && document.documentElement.classList.contains("dark")
        ? "#0f172a"
        : "#ffffff",
    pixelRatio: 2,
  });
}

export async function captureRefToDataUrl(node: HTMLElement | null): Promise<string | null> {
  if (!node) return null;
  try {
    return await nodeToPngDataUrl(node);
  } catch {
    return null;
  }
}

export async function captureElementPng(node: HTMLElement, basename: string) {
  const dataUrl = await nodeToPngDataUrl(node);
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = `${sanitizeBase(basename)}.png`;
  a.click();
}

function matrixRowLabel(row: PivotRegionRow): string {
  const rec = row as Record<string, unknown>;
  if (typeof rec.name === "string" && rec.name.length) return rec.name;
  if (typeof rec.métier === "string") return String(rec.métier);
  if (typeof rec.nature === "string") return String(rec.nature);
  return "";
}

export function exportMatrixExcel(
  title: string,
  labelHeader: string,
  rows: PivotRegionRow[],
  regionCols: string[],
  basename: string,
) {
  const wb = XLSX.utils.book_new();
  const json = rows.map((row) => {
    const o: Record<string, string | number> = { [labelHeader]: matrixRowLabel(row) };
    for (const s of regionCols) o[s] = numericValue(row, s);
    const tot = regionCols.reduce((acc, s) => acc + numericValue(row, s), 0);
    o.Total = tot;
    return o;
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(json), title.slice(0, 28));
  XLSX.writeFile(wb, `${sanitizeBase(basename)}.xlsx`);
}

export async function exportMatrixPdf(
  title: string,
  labelHeader: string,
  rows: PivotRegionRow[],
  regionCols: string[],
  basename: string,
) {
  const logo = await resolveReportLogo(null);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 280, 16, 2, 2, "F");
  if (logo) {
    try {
      doc.addImage(logo, "PNG", 10, 9, 16, 14);
    } catch {}
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(13);
  doc.text(title, 30, 17);
  doc.setFontSize(8);
  doc.text(reportTimestamp(), 200, 17);

  const head = [labelHeader, ...regionCols, "Total"];
  const body = rows.map((row) => {
    const nums = regionCols.map((s) => numericValue(row, s));
    return [matrixRowLabel(row), ...nums.map(String), String(nums.reduce((a, b) => a + b, 0))];
  });

  autoTable(doc, {
    head: [head],
    body,
    startY: 30,
    styles: { fontSize: 7 },
    headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255] },
    didParseCell(data) {
      const d = data as {
        section: string;
        column: { index: number };
        row: { raw?: unknown[] };
        cell: { styles: { textColor?: number[] } };
      };
      if (d.section === "body" && d.column.index === 0) {
        const lbl = String(d.row.raw?.[0] ?? "");
        d.cell.styles.textColor = hexToRgbTuple(getResultColor(lbl));
      }
    },
  });

  doc.save(`${sanitizeBase(basename)}.pdf`);
}

export async function exportMatrixPptx(
  title: string,
  labelHeader: string,
  rows: PivotRegionRow[],
  regionCols: string[],
  basename: string,
) {
  const logo = await resolveReportLogo(null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.background = { color: "fff7ed" };
  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: 0.65, fill: { color: "f97316" }, line: { color: "f97316", pt: 0 } });
  if (logo) {
    try {
      slide.addImage({ data: logo, x: 0.35, y: 0.1, w: 0.65, h: 0.48 });
    } catch {}
  }
  slide.addText(title, { x: 1.15, y: 0.18, fontSize: 16, bold: true, color: "FFFFFF" });
  slide.addText(reportTimestamp(), { x: 9.5, y: 0.22, fontSize: 8, color: "FFEDD5", align: "right" });

  const head = [{ text: labelHeader, options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } }, ...regionCols.map((c) => ({ text: c, options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } })), { text: "Total", options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } }];
  const tableRows = [
    head,
    ...rows.slice(0, 24).map((row) => {
      const nums = regionCols.map((s) => String(numericValue(row, s)));
      const tot = regionCols.reduce((a, s) => a + numericValue(row, s), 0);
      return [{ text: matrixRowLabel(row) }, ...nums.map((t) => ({ text: t })), { text: String(tot) }];
    }),
  ];
  slide.addTable(tableRows as never, { x: 0.4, y: 0.85, w: 12.5, fontSize: 8, colW: [2.2, ...regionCols.map(() => 1.1), 1.1] });

  await pptx.writeFile({ fileName: `${sanitizeBase(basename)}.pptx` });
}

export function exportTeleOpExcel(ops: OperatorRankRow[], metricKeys: string[], basename: string) {
  const defs = [
    { key: "name", label: "Téléopérateur" },
    { key: "volume", label: "Volume" },
    { key: "abandons", label: "Appels abandonnés" },
    { key: "appelsDécrochésInterrompus", label: "Appels décrochés interrompus" },
    { key: "informés", label: "Clients informés" },
    { key: "tickets", label: "Tickets transmis" },
  ];
  const cols = defs.filter((d) => d.key === "name" || metricKeys.includes(d.key));
  const wb = XLSX.utils.book_new();
  const json = ops.map((o) => {
    const r: Record<string, string | number> = {};
    for (const c of cols) {
      r[c.label] = (o as Record<string, string | number>)[c.key];
    }
    return r;
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(json), "Téléopérateurs");
  XLSX.writeFile(wb, `${sanitizeBase(basename)}.xlsx`);
}

export async function exportTeleOpPdf(ops: OperatorRankRow[], metricKeys: string[], basename: string) {
  const logo = await resolveReportLogo(null);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 280, 14, 2, 2, "F");
  if (logo) {
    try {
      doc.addImage(logo, "PNG", 10, 9, 14, 12);
    } catch {}
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(12);
  doc.text("Statistiques téléopérateurs", 28, 16);

  const head = ["Téléopérateur", ...metricKeys.map((k) => labelForTeleKey(k))];
  const body = ops.map((o) => [
    o.name,
    ...metricKeys.map((k) => String(numericValue(o, k))),
  ]);

  autoTable(doc, {
    head: [head],
    body,
    startY: 28,
    styles: { fontSize: 8 },
    headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255] },
    didParseCell(data) {
      const d = data as { section: string; column: { index: number }; cell: { styles: { textColor?: number[] } } };
      if (d.section !== "body") return;
      const mk = metricKeys[d.column.index - 1];
      if (!mk) return;
      const color = colorForTeleKey(mk);
      if (color) d.cell.styles.textColor = hexToRgbTuple(color);
    },
  });
  doc.save(`${sanitizeBase(basename)}.pdf`);
}

function labelForTeleKey(k: string) {
  const m: Record<string, string> = {
    volume: "Volume",
    abandons: "Appels abandonnés",
    appelsDécrochésInterrompus: "Appels décrochés interrompus",
    informés: "Clients informés",
    tickets: "Tickets transmis",
  };
  return m[k] ?? k;
}

function colorForTeleKey(k: string) {
  const m: Record<string, string> = {
    volume: AGGREGATE_VOLUME_BAR,
    abandons: RESULT_COLORS["Appels abandonnés"],
    appelsDécrochésInterrompus: RESULT_COLORS["Appels décrochés interrompus"],
    informés: RESULT_COLORS["Clients informés"],
    tickets: RESULT_COLORS["Tickets transmis"],
  };
  return m[k];
}

export async function exportTeleOpPptx(ops: OperatorRankRow[], metricKeys: string[], basename: string) {
  const logo = await resolveReportLogo(null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.background = { color: "fff7ed" };
  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: 0.65, fill: { color: "f97316" }, line: { color: "f97316", pt: 0 } });
  if (logo) {
    try {
      slide.addImage({ data: logo, x: 0.35, y: 0.1, w: 0.65, h: 0.48 });
    } catch {}
  }
  slide.addText("Statistiques téléopérateurs", { x: 1.15, y: 0.2, fontSize: 16, bold: true, color: "FFFFFF" });

  const head = [
    { text: "Téléopérateur", options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } },
    ...metricKeys.map((k) => ({
      text: labelForTeleKey(k),
      options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" },
    })),
  ];
  const body = ops.slice(0, 18).map((o) => [
    { text: o.name },
    ...metricKeys.map((k) => {
      const hex = (colorForTeleKey(k) ?? "#334155").replace("#", "");
      return {
        text: String(numericValue(o, k)),
        options: { color: hex },
      };
    }),
  ]);
  slide.addTable([head, ...body] as never, { x: 0.4, y: 0.85, w: 12.5, fontSize: 8 });

  await pptx.writeFile({ fileName: `${sanitizeBase(basename)}.pptx` });
}

/** Aperçu brut cockpit : colonnes visibles uniquement, même ordre que `keys`. */
export function exportRawPreviewExcel(rows: CrcRow[], keys: RawColumnKey[], headers: string[], basename: string) {
  const wb = XLSX.utils.book_new();
  const aoa: (string | number)[][] = [headers, ...rows.map((r) => keys.map((k) => formatRawCellForExport(r, k)))];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Aperçu brut");
  XLSX.writeFile(wb, `${sanitizeBase(basename)}.xlsx`);
}

export async function exportRawPreviewPdf(
  rows: CrcRow[],
  keys: RawColumnKey[],
  headers: string[],
  basename: string,
) {
  const logo = await resolveReportLogo(null);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 280, 14, 2, 2, "F");
  if (logo) {
    try {
      doc.addImage(logo, "PNG", 10, 8, 14, 12);
    } catch {}
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(12);
  doc.text("Grille brute (extrait)", 28, 16);
  doc.setFontSize(8);
  doc.text(reportTimestamp(), 220, 16);

  const body = rows.map((r) => keys.map((k) => formatRawCellForExport(r, k)));
  autoTable(doc, {
    head: [headers],
    body,
    startY: 26,
    styles: { fontSize: 7 },
    headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255] },
  });
  doc.save(`${sanitizeBase(basename)}.pdf`);
}

export async function exportRawPreviewPptx(
  rows: CrcRow[],
  keys: RawColumnKey[],
  headers: string[],
  basename: string,
) {
  const logo = await resolveReportLogo(null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.background = { color: "fff7ed" };
  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: 0.65, fill: { color: "f97316" }, line: { color: "f97316", pt: 0 } });
  if (logo) {
    try {
      slide.addImage({ data: logo, x: 0.35, y: 0.1, w: 0.65, h: 0.48 });
    } catch {}
  }
  slide.addText("Grille brute (extrait)", { x: 1.15, y: 0.2, fontSize: 16, bold: true, color: "FFFFFF" });

  const head = headers.map((h) => ({
    text: h,
    options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" },
  }));
  const body = rows.slice(0, 22).map((r) => keys.map((k) => ({ text: formatRawCellForExport(r, k) })));
  slide.addTable([head, ...body] as never, { x: 0.35, y: 0.85, w: 12.5, fontSize: 7 });

  await pptx.writeFile({ fileName: `${sanitizeBase(basename)}.pptx` });
}

export type RegionResultChartRow = { résultat: string; volume: number };

export function getRegionResultChartRows(rows: CrcRow[], region: CanonicalRegion): RegionResultChartRow[] {
  const a = regionSlice(rows, region);
  return a.résultat.slice(0, 8).map(([name, cnt]) => ({ résultat: name, volume: cnt }));
}

export async function exportRegionResultBarPdf(regionTitle: string, chartRows: RegionResultChartRow[], basename: string) {
  const logo = await resolveReportLogo(null);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 280, 14, 2, 2, "F");
  if (logo) {
    try {
      doc.addImage(logo, "PNG", 10, 8, 14, 12);
    } catch {}
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(12);
  doc.text(regionTitle, 28, 16);

  autoTable(doc, {
    head: [["Résultat (agrégé)", "Volume"]],
    body: chartRows.map((r) => [r.résultat, String(r.volume)]),
    startY: 26,
    styles: { fontSize: 8 },
    headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255] },
    didParseCell(data) {
      const d = data as { section: string; column: { index: number }; cell: { text?: unknown[]; styles: { textColor?: number[] } } };
      if (d.section !== "body" || d.column.index !== 0) return;
      const lbl = String(d.cell.text?.[0] ?? "");
      d.cell.styles.textColor = hexToRgbTuple(getResultColor(lbl));
    },
  });
  doc.save(`${sanitizeBase(basename)}.pdf`);
}

export async function exportRegionResultBarPptx(
  regionTitle: string,
  chartRows: RegionResultChartRow[],
  basename: string,
) {
  const logo = await resolveReportLogo(null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.background = { color: "fff7ed" };
  slide.addShape("rect", { x: 0, y: 0, w: "100%", h: 0.65, fill: { color: "f97316" }, line: { color: "f97316", pt: 0 } });
  if (logo) {
    try {
      slide.addImage({ data: logo, x: 0.35, y: 0.1, w: 0.65, h: 0.48 });
    } catch {}
  }
  slide.addText(regionTitle, { x: 1.15, y: 0.18, fontSize: 16, bold: true, color: "FFFFFF" });
  slide.addText(reportTimestamp(), { x: 9.5, y: 0.22, fontSize: 8, color: "FFEDD5", align: "right" });

  const labels = chartRows.map((r) => r.résultat);
  const values = chartRows.map((r) => r.volume);
  addResultDistributionBarChart(slide, labels, values, { x: 0.5, y: 1.05, w: 12.3, h: 3.4 });

  const head = [
    { text: "Résultat", options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } },
    { text: "Volume", options: { bold: true, fill: { color: "334155" }, color: "FFFFFF" } },
  ];
  const body = chartRows.map((r) => [
    { text: r.résultat, options: { color: hexForPptx(getResultColor(r.résultat)) } },
    { text: String(r.volume) },
  ]);
  slide.addTable([head, ...body] as never, { x: 0.5, y: 4.65, w: 12.0, fontSize: 9 });

  await pptx.writeFile({ fileName: `${sanitizeBase(basename)}.pptx` });
}

export function exportRegionResultBarExcel(regionTitle: string, chartRows: RegionResultChartRow[], basename: string) {
  const wb = XLSX.utils.book_new();
  const json = chartRows.map((r) => ({ Résultat: r.résultat, Volume: r.volume }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(json), regionTitle.slice(0, 28));
  XLSX.writeFile(wb, `${sanitizeBase(basename)}.xlsx`);
}
