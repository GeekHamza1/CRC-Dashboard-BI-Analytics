import type PptxGenJS from "pptxgenjs";

import { AGGREGATE_VOLUME_BAR, compareResultBuckets, getResultColor, hexForPptx } from "../constants/chart-colors";
import { résultatBucket } from "../crc-analytics";
import { REGION_COLORS, REGION_ORDER, REGION_SHORT } from "../crc-constants";
import type { PivotRegionRow } from "../crc-export-helpers";

function regionHexForShort(s: string): string {
  const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === s);
  return rg ? hexForPptx(REGION_COLORS[rg]) : "94a3b8";
}

/** Stacked bar — one series per visible region, categories = row labels. */
export function addStackedPivotBarChart(
  slide: PptxGenJS.Slide,
  rows: PivotRegionRow[],
  rowLabelAccessor: (row: PivotRegionRow) => string,
  regionShorts: string[],
  box: { x: number; y: number; w: number; h: number },
) {
  const labels = rows.map((r) => rowLabelAccessor(r));
  const series = regionShorts.map((s) => ({
    name: s,
    labels,
    values: rows.map((row) => Number((row as Record<string, number>)[s] ?? 0)),
  }));
  slide.addChart("bar", series, {
    ...box,
    barGrouping: "stacked",
    showLegend: true,
    chartColors: regionShorts.map((s) => regionHexForShort(s)),
  });
}

export function addRegionBarChart(
  slide: PptxGenJS.Slide,
  labels: string[],
  values: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "bar",
    [
      {
        name: "Volume",
        labels,
        values,
      },
    ],
    {
      ...box,
      showLegend: false,
      chartColors: labels.map((lib) => {
        const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === lib);
        return rg ? hexForPptx(REGION_COLORS[rg]) : "94a3b8";
      }),
    },
  );
}

export function addRegionDoughnutChart(
  slide: PptxGenJS.Slide,
  labels: string[],
  values: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "doughnut",
    [
      {
        name: "Régions",
        labels,
        values,
      },
    ],
    {
      ...box,
      showLegend: true,
      holeSize: 55,
      chartColors: labels.map((lib) => {
        const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === lib);
        return rg ? hexForPptx(REGION_COLORS[rg]) : "94a3b8";
      }),
    },
  );
}

export function addStatusPieChart(
  slide: PptxGenJS.Slide,
  names: string[],
  values: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "pie",
    [
      {
        name: "Résultat",
        labels: names,
        values,
      },
    ],
    {
      ...box,
      showLegend: true,
      chartColors: names.map((n) => hexForPptx(getResultColor(n))),
    },
  );
}

export function addLineTotalsChart(
  slide: PptxGenJS.Slide,
  dayLabels: string[],
  totals: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "line",
    [
      {
        name: "Total journalier",
        labels: dayLabels,
        values: totals,
      },
    ],
    {
      ...box,
      lineSize: 2,
      showLegend: false,
      chartColors: [hexForPptx(AGGREGATE_VOLUME_BAR)],
    },
  );
}

export function addMonthlyTotalsBarChart(
  slide: PptxGenJS.Slide,
  monthLabels: string[],
  totals: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "bar",
    [
      {
        name: "Total mensuel",
        labels: monthLabels,
        values: totals,
      },
    ],
    {
      ...box,
      showLegend: false,
      chartColors: [hexForPptx(AGGREGATE_VOLUME_BAR)],
    },
  );
}

export function addStackedMonthlyByRegionChart(
  slide: PptxGenJS.Slide,
  series: { name: string; labels: string[]; values: number[] }[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart("bar", series, {
    ...box,
    barGrouping: "stacked",
    showLegend: true,
    chartColors: series.map((s) => {
      const rg = REGION_ORDER.find((r) => REGION_SHORT[r] === s.name);
      return rg ? hexForPptx(REGION_COLORS[rg]) : "94a3b8";
    }),
  });
}

export function addAreaTotalsChart(
  slide: PptxGenJS.Slide,
  dayLabels: string[],
  totals: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "area",
    [
      {
        name: "Total",
        labels: dayLabels,
        values: totals,
      },
    ],
    {
      ...box,
      showLegend: false,
      chartColors: [hexForPptx(AGGREGATE_VOLUME_BAR)],
    },
  );
}

export function addTeleopVolumeBarChart(
  slide: PptxGenJS.Slide,
  names: string[],
  volumes: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  slide.addChart(
    "bar",
    [
      {
        name: "Volume",
        labels: names,
        values: volumes,
      },
    ],
    {
      ...box,
      showLegend: false,
      chartColors: [hexForPptx(AGGREGATE_VOLUME_BAR)],
    },
  );
}

/**
 * Carte région « Résultat » — barres natives, couleurs = familles CRC (getResultColor),
 * aligné sur le dashboard (regionSlice).
 */
export function addResultDistributionBarChart(
  slide: PptxGenJS.Slide,
  labels: string[],
  values: number[],
  box: { x: number; y: number; w: number; h: number },
) {
  const chartColors = labels.map((lbl) => hexForPptx(getResultColor(lbl)));
  slide.addChart(
    "bar",
    [
      {
        name: "Interactions",
        labels,
        values,
      },
    ],
    {
      ...box,
      showLegend: false,
      chartColors,
    },
  );
}

/** Build résultat bucket pie data (same logic as dashboard). */
export function buildRésultatPieSeries(rows: { résultat: string }[]) {
  const buckets = rows.map((r) => résultatBucket(r.résultat));
  const uniq = [...new Set(buckets)].sort(compareResultBuckets);
  const names = uniq.filter((n) => buckets.filter((x) => x === n).length > 0);
  const values = names.map((name) => buckets.filter((x) => x === name).length);
  return { names, values };
}
