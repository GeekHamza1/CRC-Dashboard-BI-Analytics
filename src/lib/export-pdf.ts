"use client";

import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import {
  dailySeries,
  globalKpis,
  monthlySeries,
  operatorRanking,
  pivotMétierParRégion,
  pivotNatureParRégion,
  pivotRésultatParRégion,
} from "./crc-analytics";
import {
  AGGREGATE_VOLUME_BAR,
  compareResultBuckets,
  getResultColor,
  RESULT_COLORS,
  hexToRgbTuple,
} from "./constants/chart-colors";
import { REGION_COLORS, REGION_ORDER, REGION_SHORT } from "./crc-constants";
import type { CrcExportColumnVisibility } from "./crc-export-helpers";
import { activeRegionShorts, activeTeleOpKeys } from "./crc-export-helpers";
import type { CrcReportConfig } from "./crc-report-config";
import { mergeExportPdf } from "./crc-report-config";
import type { CrcRow } from "./crc-types";
import { reportTimestamp, resolveReportLogo } from "./export-branding";

type PdfPoint = {
  label: string;
  value: number;
  color: string;
};

function drawHeaderBand(
  doc: jsPDF,
  title: string,
  subtitle: string,
  logoDataUrl: string | null,
) {
  doc.setFillColor(249, 115, 22);
  doc.roundedRect(8, 8, 194, 22, 3, 3, "F");
  if (logoDataUrl) {
    try {
      doc.addImage(logoDataUrl, "PNG", 11, 9.2, 22, 18);
    } catch {
      /** image optional */
    }
  }
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(15);
  doc.text(title, 38, 16.5);
  doc.setFontSize(9);
  doc.text(subtitle, 38, 22);
}

function drawFooter(doc: jsPDF, pageTitle: string) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  doc.setDrawColor(226, 232, 240);
  doc.line(10, pageHeight - 11, pageWidth - 10, pageHeight - 11);
  doc.setFontSize(8.5);
  doc.setTextColor(100, 116, 139);
  doc.text(pageTitle, 10, pageHeight - 6.3);
  doc.text(
    `Généré le ${reportTimestamp()}`,
    pageWidth - 10,
    pageHeight - 6.3,
    { align: "right" },
  );
}

function drawKpiCard(
  doc: jsPDF,
  x: number,
  y: number,
  w: number,
  h: number,
  label: string,
  value: string | number,
  accentHex: string,
) {
  const [r, g, b] = hexToRgbTuple(accentHex);
  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(241, 245, 249);
  doc.roundedRect(x, y, w, h, 3, 3, "FD");
  doc.setFillColor(r, g, b);
  doc.roundedRect(x, y, 1.8, h, 1.5, 1.5, "F");
  doc.setTextColor(71, 85, 105);
  doc.setFontSize(8.8);
  doc.text(label, x + 3.5, y + 5);
  doc.setTextColor(15, 23, 42);
  doc.setFontSize(15);
  doc.text(String(value), x + 3.5, y + 11.5);
}

function drawBarChart(
  doc: jsPDF,
  x: number,
  y: number,
  w: number,
  h: number,
  title: string,
  points: PdfPoint[],
) {
  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(226, 232, 240);
  doc.roundedRect(x, y, w, h, 3, 3, "FD");
  doc.setTextColor(30, 41, 59);
  doc.setFontSize(10);
  doc.text(title, x + 2.5, y + 5);
  const chartTop = y + 9;
  const chartBottom = y + h - 8;
  const max = Math.max(1, ...points.map((p) => p.value));
  const barZoneW = w - 12;
  const barW = Math.max(6, barZoneW / Math.max(1, points.length) - 3);
  points.forEach((p, idx) => {
    const px = x + 5 + idx * (barW + 3);
    const barH = ((chartBottom - chartTop) * p.value) / max;
    const py = chartBottom - barH;
    const [r, g, b] = hexToRgbTuple(p.color);
    doc.setFillColor(r, g, b);
    doc.roundedRect(px, py, barW, barH, 1.2, 1.2, "F");
    doc.setTextColor(71, 85, 105);
    doc.setFontSize(7);
    doc.text(String(p.value), px + barW / 2, py - 1.5, { align: "center" });
    doc.text(p.label.slice(0, 11), px + barW / 2, chartBottom + 4.2, {
      align: "center",
    });
  });
}

function drawLineChart(
  doc: jsPDF,
  x: number,
  y: number,
  w: number,
  h: number,
  title: string,
  values: number[],
  colorHex: string,
) {
  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(226, 232, 240);
  doc.roundedRect(x, y, w, h, 3, 3, "FD");
  doc.setTextColor(30, 41, 59);
  doc.setFontSize(10);
  doc.text(title, x + 2.5, y + 5);
  if (!values.length) return;
  const innerX = x + 6;
  const innerY = y + 10;
  const innerW = w - 10;
  const innerH = h - 16;
  const max = Math.max(1, ...values);
  const [r, g, b] = hexToRgbTuple(colorHex);
  doc.setDrawColor(r, g, b);
  doc.setLineWidth(0.9);
  values.forEach((v, i) => {
    const px = innerX + (innerW * i) / Math.max(1, values.length - 1);
    const py = innerY + innerH - (innerH * v) / max;
    if (i > 0) {
      const ppx = innerX + (innerW * (i - 1)) / Math.max(1, values.length - 1);
      const ppy = innerY + innerH - (innerH * values[i - 1]) / max;
      doc.line(ppx, ppy, px, py);
    }
    doc.setFillColor(r, g, b);
    doc.circle(px, py, 0.9, "F");
  });
}

function lastTableY(doc: jsPDF): number {
  return (doc as unknown as { lastAutoTable?: { finalY: number } })
    .lastAutoTable?.finalY ?? 20;
}

export async function exportCrcPdf(
  rows: CrcRow[],
  title = "Département Clientèle et Suivi des Performances CRC",
  basename = "crc_axilus_export",
  opts?: {
    subtitle?: string;
    logoOverride?: string | null;
    pdf?: Partial<CrcReportConfig["exportPdf"]>;
    /** Colonnes visibles cockpit — même état que le tableau de bord */
    columnVisibility?: CrcExportColumnVisibility;
    /** Captures PNG (data URL) des graphiques associés aux widgets */
    widgetChartImages?: Partial<Record<"pivotResult" | "pivotMetier" | "pivotNature" | "teleOp", string>>;
  },
) {
  const pdfCfg = mergeExportPdf(opts?.pdf);
  const logoDataUrl = await resolveReportLogo(opts?.logoOverride ?? null);
  const margin = 10;
  const cv = opts?.columnVisibility;
  const visPivotResult = activeRegionShorts(cv?.pivotResultRegions);
  const visPivotMetier = activeRegionShorts(cv?.pivotMetierRegions);
  const visPivotNature = activeRegionShorts(cv?.pivotNatureRegions);
  const teleKeys = activeTeleOpKeys(cv?.teleOpMetrics);
  const subtitle =
    opts?.subtitle?.trim() ||
    `CRC opérationnel • ${rows.length} interactions Axilus (libellés Résultat bruts agrégés par familles)`;

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  drawHeaderBand(doc, title, subtitle, logoDataUrl);

  const k = globalKpis(rows);
  const resultPivot = pivotRésultatParRégion(rows);
  const daySeries = dailySeries(rows).slice(-20);
  const monthSeries = monthlySeries(rows).slice(-12);
  const ops = operatorRanking(rows).slice(0, 12);

  let cursor = 34;
  if (pdfCfg.kpisCards) {
    drawKpiCard(
      doc,
      10,
      34,
      37,
      15,
      "Total réclamations",
      k.réclamations,
      RESULT_COLORS["Tickets transmis"],
    );
    drawKpiCard(doc, 50, 34, 37, 15, "Appels abandonnés", k.appelsAbandonnés, RESULT_COLORS["Appels abandonnés"]);
    drawKpiCard(
      doc,
      90,
      34,
      37,
      15,
      "Clients informés",
      k.clientsInformés,
      RESULT_COLORS["Clients informés"],
    );
    drawKpiCard(
      doc,
      130,
      34,
      37,
      15,
      "Tickets transmis",
      k.ticketsTransmis,
      RESULT_COLORS["Tickets transmis"],
    );
    drawKpiCard(doc, 170, 34, 32, 15, "Total appels", k.totalAppels, "#f97316");
    cursor = 52;
  }

  const statusCounts = resultPivot
    .map((r) => ({
      label: r.name,
      value: visPivotResult.reduce(
        (s, key) => s + Number((r as Record<string, number>)[key] ?? 0),
        0,
      ),
      color: getResultColor(r.name),
    }))
    .sort((a, b) => compareResultBuckets(a.label, b.label));

  const chartTop = pdfCfg.kpisCards ? 53 : cursor + 4;
  if (pdfCfg.summaryCharts) {
    drawBarChart(doc, 10, chartTop, 95, 50, "Résultat global", statusCounts.slice(0, 5));
    drawBarChart(
      doc,
      107,
      chartTop,
      95,
      50,
      "Volumes par région",
      REGION_ORDER.filter((rg) => visPivotResult.includes(REGION_SHORT[rg])).map((rg) => ({
        label: REGION_SHORT[rg],
        value: k.appelsParRégion.get(rg) ?? 0,
        color: REGION_COLORS[rg],
      })),
    );
    cursor = chartTop + 54;
  } else if (pdfCfg.kpisCards) {
    cursor = chartTop + 49;
  } else if (!pdfCfg.kpisCards) {
    cursor = Math.max(cursor, 34);
  }

  let page1HadContent = !!(pdfCfg.kpisCards || pdfCfg.summaryCharts);
  if (pdfCfg.pivotResultTable) {
    autoTable(doc, {
      head: [["Résultat", ...visPivotResult, "Total"]],
      body: resultPivot.map((row) => {
        const vals = visPivotResult.map((s) => Number((row as Record<string, number>)[s] ?? 0));
        return [row.name, ...vals.map(String), String(vals.reduce((a, b) => a + b, 0))];
      }),
      startY: cursor + 6,
      styles: { fontSize: 8 },
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
    page1HadContent = true;
  }

  if (opts?.widgetChartImages?.pivotResult && pdfCfg.pivotResultTable) {
    try {
      const yImg = (doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable?.finalY ?? 120;
      doc.addImage(opts.widgetChartImages.pivotResult, "PNG", margin, yImg + 4, 180, 72);
    } catch {
      /** ignore bad image */
    }
  }

  drawFooter(doc, page1HadContent || pdfCfg.pivotResultTable ? "Page 1 — Synthèse globale" : "CRC export");

  if (pdfCfg.trendsCharts || pdfCfg.metierNatureTables) {
    doc.addPage();
    drawHeaderBand(doc, "Récapitulatif opérationnel", "Évolution temporelle et traitement CRC", logoDataUrl);
    let tCursor = 36;
    if (pdfCfg.trendsCharts) {
      drawLineChart(
        doc,
        margin,
        tCursor,
        92,
        58,
        "Évolution journalière (total)",
        daySeries.map((d) => Number(d.total)),
        AGGREGATE_VOLUME_BAR,
      );
      drawLineChart(
        doc,
        106,
        tCursor,
        96,
        58,
        "Évolution mensuelle (total)",
        monthSeries.map((m) => Number(m.total)),
        "#d97706",
      );
      tCursor = 98;
    } else if (pdfCfg.metierNatureTables) {
      tCursor = 42;
    }

    if (pdfCfg.metierNatureTables) {
      const métierTop = pivotMétierParRégion(rows).slice(0, 8);
      autoTable(doc, {
        head: [["Métier", ...visPivotMetier]],
        body: métierTop.map((r) => [
          String(r.métier),
          ...visPivotMetier.map((reg) => String(r[reg] ?? 0)),
        ]),
        startY: tCursor,
        styles: { fontSize: 7.5 },
        headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255] },
      });

      let yGap = lastTableY(doc);
      if (opts?.widgetChartImages?.pivotMetier) {
        try {
          doc.addImage(opts.widgetChartImages.pivotMetier, "PNG", margin, yGap + 4, 180, 52);
          yGap += 4 + 52;
        } catch {
          /** ignore */
        }
      }

      const natureTop = pivotNatureParRégion(rows).slice(0, 8);
      autoTable(doc, {
        head: [["Nature de Réclamation", ...visPivotNature]],
        body: natureTop.map((r) => [
          String(r.nature),
          ...visPivotNature.map((reg) => String(r[reg] ?? 0)),
        ]),
        startY: yGap + 6,
        styles: { fontSize: 7.3 },
        headStyles: { fillColor: [180, 83, 9], textColor: [255, 255, 255] },
      });

      if (opts?.widgetChartImages?.pivotNature) {
        try {
          const yn = lastTableY(doc) + 4;
          doc.addImage(opts.widgetChartImages.pivotNature, "PNG", margin, yn, 180, 50);
        } catch {
          /** ignore */
        }
      }
    }
    drawFooter(doc, "Métiers & réclamations");
  }

  if (pdfCfg.teleopPage) {
    doc.addPage();
    drawHeaderBand(
      doc,
      "Statistiques des Téléopérateurs",
      "Colonnes depuis le champ Résultat brut Axilus — Abandon / Appel abandonné / Clients informés / Tickets",
      logoDataUrl,
    );

    drawBarChart(
      doc,
      margin,
      36,
      192,
      55,
      "Top 10 volumes téléopérateur",
      ops.slice(0, 10).map((o) => ({
        label: o.name,
        value: o.volume,
        color: AGGREGATE_VOLUME_BAR,
      })),
    );

    const teleHead = [
      "Téléopérateur",
      ...teleKeys.map((key) => {
        const labels: Record<string, string> = {
          volume: "Volume",
          abandons: "Appels abandonnés",
          appelsDécrochésInterrompus: "Appels décrochés interrompus",
          informés: "Clients informés",
          tickets: "Tickets transmis",
        };
        return labels[key] ?? key;
      }),
    ];
    const teleColorByCol: (string | null)[] = [
      null,
      ...teleKeys.map((key) => {
        const m: Record<string, string> = {
          volume: AGGREGATE_VOLUME_BAR,
          abandons: RESULT_COLORS["Appels abandonnés"],
          appelsDécrochésInterrompus: RESULT_COLORS["Appels décrochés interrompus"],
          informés: RESULT_COLORS["Clients informés"],
          tickets: RESULT_COLORS["Tickets transmis"],
        };
        return m[key] ?? null;
      }),
    ];
    autoTable(doc, {
      head: [teleHead],
      body: ops.map((o) => [
        o.name,
        ...teleKeys.map((key) => String((o as Record<string, number>)[key] ?? 0)),
      ]),
      startY: 96,
      styles: { fontSize: 8 },
      headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255] },
      didParseCell(data) {
        const d = data as {
          section: string;
          column: { index: number };
          cell: { styles: { textColor?: number[] } };
        };
        if (d.section !== "body") return;
        const hex = teleColorByCol[d.column.index];
        if (hex) d.cell.styles.textColor = hexToRgbTuple(hex);
      },
      foot: [
        [
          "* Agrégats basés sur le texte brut Résultat (Axilus)",
          ...Array(Math.max(0, teleHead.length - 1)).fill(""),
        ],
      ],
    });
    if (opts?.widgetChartImages?.teleOp) {
      try {
        const yTele = lastTableY(doc) + 4;
        doc.addImage(opts.widgetChartImages.teleOp, "PNG", margin, yTele, 180, 55);
      } catch {
        /** ignore */
      }
    }
    drawFooter(doc, "Classement téléopérateurs");
  }

  const fn = `${basename.replace(/[^a-zA-Z0-9_-]+/g, "_")}.pdf`;
  doc.save(fn);
}
