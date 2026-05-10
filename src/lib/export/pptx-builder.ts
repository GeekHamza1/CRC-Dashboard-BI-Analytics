import PptxGenJS from "pptxgenjs";

import {
  dailySeries,
  globalKpis,
  monthlySeries,
  operatorRanking,
  pivotMétierParRégion,
  pivotNatureParRégion,
  pivotRésultatParRégion,
} from "../crc-analytics";
import { RESULT_COLORS, hexForPptx } from "../constants/chart-colors";
import { REGION_COLORS, REGION_ORDER, REGION_SHORT } from "../crc-constants";
import type { CrcExportColumnVisibility } from "../crc-export-helpers";
import { activeRegionShorts, activeTeleOpKeys, numericValue, type PivotRegionRow } from "../crc-export-helpers";
import { RAW_COLUMN_KEYS, mergeExportPptx } from "../crc-report-config";
import type { CrcRow } from "../crc-types";
import { resolveReportLogo } from "../export-branding";

import type { CrcPptxDashboardSnapshot, CrcPptxExportOptions } from "./crc-pptx-types";
import {
  addAreaTotalsChart,
  addLineTotalsChart,
  addMonthlyTotalsBarChart,
  addRegionBarChart,
  addRegionDoughnutChart,
  addResultDistributionBarChart,
  addStackedMonthlyByRegionChart,
  addStackedPivotBarChart,
  addStatusPieChart,
  addTeleopVolumeBarChart,
  buildRésultatPieSeries,
} from "./pptx-charts";
import { addKpiCardGrid, buildKpiItems } from "./pptx-kpis";
import { addSlideHeader } from "./pptx-layout";
import { addPivotTable, addRawPreviewTable, teleopTableRows } from "./pptx-tables";
import { PPTX_CRC } from "./pptx-theme";
import { buildCrcPptxWidgetPlan } from "./pptx-widgets";
import { buildCrcPptxWidgetRegistry } from "./widget-registry";

export async function buildAndSaveCrcPptx(
  rows: CrcRow[],
  title: string,
  basename: string,
  opts: CrcPptxExportOptions,
) {
  const dash = opts.dashboard as CrcPptxDashboardSnapshot;
  const slideCfg = mergeExportPptx({ ...dash.exportPptx, ...(opts.pptx ?? {}) });
  void buildCrcPptxWidgetPlan(dash);
  const widgetRegistry = buildCrcPptxWidgetRegistry(rows, dash);

  const cv: CrcExportColumnVisibility | undefined = opts.columnVisibility;
  const visR = activeRegionShorts(cv?.pivotResultRegions);
  const visM = activeRegionShorts(cv?.pivotMetierRegions);
  const visN = activeRegionShorts(cv?.pivotNatureRegions);
  const teleK = activeTeleOpKeys(cv?.teleOpMetrics);

  const logoDataUrl = await resolveReportLogo(opts.logoOverride ?? null);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "CRC Analytics";

  const k = globalKpis(rows);
  const pivR = pivotRésultatParRégion(rows);
  const pivM = pivotMétierParRégion(rows);
  const pivN = pivotNatureParRégion(rows);
  const dly = dailySeries(rows).slice(-24);
  const mon = monthlySeries(rows).slice(-14);
  const topOps = operatorRanking(rows).slice(0, 12);

  const subtitle =
    opts.subtitle?.trim() || "Statistiques des réclamations CRC — export natif PowerPoint (objets éditables)";

  if (slideCfg.cover) {
    const sd = pptx.addSlide();
    addSlideHeader(sd, "Département Clientèle et Suivi des Performances CRC", subtitle, logoDataUrl);
    sd.addText(title, {
      x: 0.8,
      y: 1.35,
      w: 11.8,
      h: 0.6,
      fontSize: 28,
      bold: true,
      color: PPTX_CRC.navy,
    });
    sd.addText(`${rows.length.toLocaleString("fr-FR")} interactions (jeu filtré exporté)`, {
      x: 0.8,
      y: 2.35,
      w: 10,
      h: 0.35,
      fontSize: 13,
      color: PPTX_CRC.slate600,
    });
  }

  if (slideCfg.kpi) {
    const kp = pptx.addSlide();
    addSlideHeader(kp, "Récapitulatif global CRC", "KPI — cartes natives PowerPoint", logoDataUrl);
    const items = buildKpiItems(rows, dash.kpis, k);
    if (items.length) {
      addKpiCardGrid(kp, items, { x: 0.45, y: 1.0 }, 4);
    }
  }

  const showRegionalSlide =
    rows.length > 0 && (dash.charts.geoBars || dash.charts.geoDonut) && (slideCfg.kpi || slideCfg.results);

  if (showRegionalSlide) {
    const rg = pptx.addSlide();
    addSlideHeader(rg, "Analytique régionale", "Volumes par région canon (données filtrées)", logoDataUrl);
    const labels = REGION_ORDER.map((r) => REGION_SHORT[r]);
    const values = REGION_ORDER.map((r) => k.appelsParRégion.get(r) ?? 0);
    if (dash.charts.geoBars) {
      addRegionBarChart(rg, labels, values, { x: 0.5, y: 1.05, w: 7.2, h: 3.2 });
    }
    if (dash.charts.geoDonut) {
      addRegionDoughnutChart(rg, labels, values, {
        x: dash.charts.geoBars ? 7.9 : 2.5,
        y: 1.05,
        w: dash.charts.geoBars ? 5.2 : 8,
        h: 3.2,
      });
    }
  }

  /** Cartes région — résultats (registry `regionCards`, natif PPT). */
  if (widgetRegistry.regionCards.length > 0) {
    const rc = pptx.addSlide();
    addSlideHeader(
      rc,
      "Régions — répartition par résultat",
      "Drâa · Laâyoune · Souss · Faux appels — graphiques éditables (même palette familles CRC)",
      logoDataUrl,
    );
    const boxes: { x: number; y: number; w: number; h: number }[] = [
      { x: 0.4, y: 1.15, w: 6.15, h: 2.35 },
      { x: 6.75, y: 1.15, w: 6.15, h: 2.35 },
      { x: 0.4, y: 3.65, w: 6.15, h: 2.35 },
      { x: 6.75, y: 3.65, w: 6.15, h: 2.35 },
    ];
    widgetRegistry.regionCards.forEach((spec, i) => {
      const box = boxes[i];
      if (!box) return;
      rc.addText(spec.widgetTitle, {
        x: box.x,
        y: box.y - 0.34,
        w: box.w,
        h: 0.28,
        fontSize: 11,
        bold: true,
        color: hexForPptx(REGION_COLORS[spec.region]),
      });
      addResultDistributionBarChart(rc, spec.chartLabels, spec.chartDataset, box);
    });
  }

  const hasResultsContent =
    slideCfg.results &&
    ((pivR.length > 0 && visR.length > 0) ||
      (dash.charts.statusPie && rows.length > 0) ||
      (dash.tables.pivotResult && pivR.length > 0) ||
      dash.charts.dailyArea ||
      dash.charts.trendLine ||
      dash.charts.monthlyBars);

  if (hasResultsContent) {
    const res = pptx.addSlide();
    addSlideHeader(
      res,
      "Résultat & dynamiques",
      "Graphiques natifs + tableau pivot (colonnes visibles cockpit)",
      logoDataUrl,
    );

    let y = 1.05;
    if (pivR.length && visR.length) {
      addStackedPivotBarChart(res, pivR as PivotRegionRow[], (row) => String(row.name), visR, {
        x: 0.45,
        y,
        w: 12.4,
        h: 2.45,
      });
      y += 2.55;
    }

    if (dash.charts.statusPie && rows.length) {
      const pie = buildRésultatPieSeries(rows);
      if (pie.names.length) {
        addStatusPieChart(res, pie.names, pie.values, { x: 0.45, y, w: 5.8, h: 2.35 });
        if (dash.tables.pivotResult && pivR.length) {
          addPivotTable(res, pivR, visR, "Résultat", { x: 6.45, y, w: 6.35, fontSize: 7 });
        }
        y += 2.45;
      } else if (dash.tables.pivotResult && pivR.length) {
        addPivotTable(res, pivR, visR, "Résultat", { x: 0.45, y, w: 12.4, fontSize: 8 });
        y += 2.5;
      }
    } else if (dash.tables.pivotResult && pivR.length) {
      addPivotTable(res, pivR, visR, "Résultat", { x: 0.45, y, w: 12.4, fontSize: 8 });
      y += 2.5;
    }

    if (dash.charts.dailyArea && dly.length) {
      addAreaTotalsChart(
        res,
        dly.map((d) => String(d.jour)),
        dly.map((d) => Number(d.total)),
        { x: 0.45, y, w: 6.0, h: 1.85 },
      );
    }
    if (dash.charts.trendLine && dly.length) {
      addLineTotalsChart(
        res,
        dly.map((d) => String(d.jour)),
        dly.map((d) => Number(d.total)),
        {
          x: dash.charts.dailyArea ? 6.55 : 0.45,
          y,
          w: dash.charts.dailyArea ? 6.3 : 12.4,
          h: 1.85,
        },
      );
    }
    if (dash.charts.dailyArea || dash.charts.trendLine) y += 1.95;

    if (dash.charts.monthlyBars && mon.length) {
      const shorts = REGION_ORDER.map((r) => REGION_SHORT[r]);
      const series = shorts.map((s) => ({
        name: s,
        labels: mon.map((m) => String(m.mois)),
        values: mon.map((m) => numericValue(m, s)),
      }));
      addStackedMonthlyByRegionChart(res, series, { x: 0.45, y, w: 12.4, h: 2.0 });
    }
  }

  if (slideCfg.metier) {
    if (dash.tables.pivotMetier && pivM.length) {
      const met = pptx.addSlide();
      addSlideHeader(met, "Métier × régions", "Histogramme empilé + tableau éditable", logoDataUrl);
      addStackedPivotBarChart(met, pivM as PivotRegionRow[], (row) => String(row.métier ?? ""), visM, {
        x: 0.45,
        y: 1.0,
        w: 12.4,
        h: 2.35,
      });
      addPivotTable(met, pivM as PivotRegionRow[], visM, "Métier", { x: 0.45, y: 3.5, w: 12.4, fontSize: 8 });
    }

    if (dash.tables.pivotNature && pivN.length) {
      const nat = pptx.addSlide();
      addSlideHeader(nat, "Nature de réclamation × régions", "Empilement + tableau", logoDataUrl);
      addStackedPivotBarChart(nat, pivN as PivotRegionRow[], (row) => String(row.nature ?? ""), visN, {
        x: 0.45,
        y: 1.0,
        w: 12.4,
        h: 2.35,
      });
      addPivotTable(nat, pivN as PivotRegionRow[], visN, "Nature de Réclamation", { x: 0.45, y: 3.5, w: 12.4, fontSize: 8 });
    }
  }

  if (slideCfg.operators && (dash.charts.teleopBars || dash.tables.teleOpStats) && topOps.length) {
    const tr = pptx.addSlide();
    addSlideHeader(tr, "Téléopérateurs", "Classement — graphique et métriques visibles", logoDataUrl);
    if (dash.charts.teleopBars) {
      addTeleopVolumeBarChart(
        tr,
        topOps.map((o) => (o.name.length > 28 ? `${o.name.slice(0, 27)}…` : o.name)),
        topOps.map((o) => o.volume),
        { x: 0.45, y: 1.0, w: 12.4, h: 2.6 },
      );
    }
    if (dash.tables.teleOpStats) {
      const rowsTable = teleopTableRows(topOps, teleK as string[]);
      tr.addTable(rowsTable as never, {
        x: 0.45,
        y: dash.charts.teleopBars ? 3.75 : 1.0,
        w: 12.1,
        colW: [3.2, ...teleK.map(() => 12.1 / Math.max(teleK.length + 1, 1) - 0.45)],
        fontSize: 8,
        border: { pt: 0.5, color: PPTX_CRC.cardBorder },
      });
    }
  }

  if (dash.tables.rawPreview && rows.length) {
    const keys = RAW_COLUMN_KEYS.filter((col) => dash.rawColumns[col]);
    if (keys.length) {
      const raw = pptx.addSlide();
      addSlideHeader(raw, "Grille brute (extrait)", "Colonnes visibles cockpit — texte éditable", logoDataUrl);
      const headers = keys.map((k) => {
        const labels: Record<string, string> = {
          date: "Date",
          resultatRaw: "Résultat (Excel)",
          teleop: "Téléopérateur",
          metier: "Métier",
          nature: "Nature",
          regionCanon: "Région canon",
          regionsource: "Région brute",
          valid: "Valide ?",
          phone: "Téléphone",
        };
        return labels[k] ?? k;
      });
      addRawPreviewTable(raw, rows, keys, headers, { x: 0.45, y: 1.0, w: 12.4, maxRows: 28 });
    }
  }

  if (slideCfg.definitions) {
    const defs = pptx.addSlide();
    addSlideHeader(defs, "Définitions CRC", "Référentiel + tendance mensuelle native", logoDataUrl);

    defs.addShape("roundRect", {
      x: 0.5,
      y: 1.0,
      w: 12.2,
      h: 2.85,
      fill: { color: "fffaf0" },
      line: { color: "fdba74", pt: 1 },
      rectRadius: 0.06,
    });

    defs.addText(
      [
        { text: "Abandon", options: { bold: true, color: hexForPptx(RESULT_COLORS["Appels abandonnés"]) } },
        { text: " → libellé Axilus « Abandon ».\n" },
        { text: "Appel abandonné", options: { bold: true, color: hexForPptx(RESULT_COLORS["Appels décrochés interrompus"]) } },
        { text: " → décrochet / interruption.\n" },
        { text: "Clients informés", options: { bold: true, color: hexForPptx(RESULT_COLORS["Clients informés"]) } },
        { text: " → information donnée.\n" },
        { text: "Tickets transmis", options: { bold: true, color: hexForPptx(RESULT_COLORS["Tickets transmis"]) } },
        { text: " → escalade dossier." },
      ],
      { x: 0.75, y: 1.2, w: 11.5, h: 2.4, fontSize: 12, color: PPTX_CRC.navy },
    );

    if (mon.length) {
      addMonthlyTotalsBarChart(
        defs,
        mon.map((m) => String(m.mois)),
        mon.map((m) => Number(m.total)),
        { x: 0.55, y: 4.05, w: 12.0, h: 1.35 },
      );
    }
  }

  await pptx.writeFile({ fileName: `${basename.replace(/\.pptx$/i, "")}.pptx` });
}
