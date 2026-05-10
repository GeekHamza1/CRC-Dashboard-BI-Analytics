"use client";

import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import PptxGenJS from "pptxgenjs";
import * as XLSX from "xlsx";
import { resolveReportLogo, reportTimestamp } from "./export-branding";

export type TableSlice = Record<string, string | number>;

export async function exportDomAsPng(node: HTMLElement, basename: string) {
  const { toPng } = await import("html-to-image");
  const img = await toPng(node, {
    cacheBust: true,
    backgroundColor:
      typeof document !== "undefined" && document.documentElement.classList.contains("dark")
        ? "#0f172a"
        : "#ffffff",
    pixelRatio: 2,
  });
  const a = document.createElement("a");
  a.href = img;
  a.download = `${sanitize(basename)}.png`;
  a.click();
}

function sanitize(s: string) {
  return s.replace(/[^\wÀ-ÖØ-öø-ÿ\-]+/gi, "_").slice(0, 160);
}

export async function exportTablePdf(
  title: string,
  subtitle: string,
  columns: string[],
  rows: TableSlice[],
  basename: string,
  logoHint?: string | null,
) {
  const logo = await resolveReportLogo(logoHint ?? null);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 280, 20, 2, 2, "F");
  if (logo) {
    try {
      doc.addImage(logo, "PNG", 11, 9, 20, 16);
    } catch {}
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(14);
  doc.text(title, 36, 16);
  doc.setFontSize(9);
  doc.text(`${subtitle} — ${reportTimestamp()}`, 36, 21.6);

  const body = rows.map((r) => columns.map((c) => String(r[c] ?? "—")));

  autoTable(doc, {
    head: [columns],
    body,
    startY: 34,
    styles: { fontSize: 7.8 },
    headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255] },
    theme: "striped",
    horizontalPageBreak: true,
    margin: { top: 32 },
  });

  doc.save(`${sanitize(basename)}.pdf`);
}

export function exportRowsExcel(columns: string[], rows: TableSlice[], sheetName: string, basename: string) {
  const wb = XLSX.utils.book_new();
  const aoa = [columns, ...rows.map((r) => columns.map((c) => r[c] ?? ""))];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheetName.slice(0, 28));
  const intro = XLSX.utils.aoa_to_sheet([
    ["Business Intelligence Export"],
    ["Généré le", reportTimestamp()],
    [],
    ["Colonnes"],
    columns,
  ]);
  XLSX.utils.book_append_sheet(wb, intro, "Meta");
  XLSX.writeFile(wb, `${sanitize(basename)}.xlsx`);
}

export async function exportTablePptx(
  title: string,
  subtitle: string,
  columns: string[],
  rows: TableSlice[],
  basename: string,
  logoHint?: string | null,
) {
  const logo = await resolveReportLogo(logoHint ?? null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  const slide = pptx.addSlide();
  slide.background = { color: "fff7ed" };

  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.75,
    fill: { color: "f97316" },
    line: { color: "f97316", pt: 0 },
  });

  if (logo) {
    try {
      slide.addImage({
        data: logo,
        x: 0.35,
        y: 0.12,
        w: 0.75,
        h: 0.55,
      });
    } catch {}
  }

  slide.addText(title, { x: 1.35, y: 0.2, fontSize: 18, bold: true, color: "FFFFFF" });
  slide.addText(`${subtitle}\n${reportTimestamp()}`, { x: 1.35, y: 0.48, fontSize: 9, color: "FFEDD5" });

  const tableRows: string[][] = [
    columns,
    ...rows.map((r) => columns.map((c) => String(r[c] ?? "—")).slice(0, columns.length)),
  ];

  slide.addTable(tableRows.slice(0, Math.min(tableRows.length, 42)), {
    x: 0.4,
    y: 1,
    w: 12.6,
    fontSize: 8,
    border: { pt: 0.4, color: "e5e7eb" },
    fill: "ffffff",
  });

  await pptx.writeFile({ fileName: `${sanitize(basename)}.pptx` });
}
