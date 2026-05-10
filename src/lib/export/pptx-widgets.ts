import type { CrcPptxDashboardSnapshot, CrcPptxWidgetDescriptor } from "./crc-pptx-types";

/**
 * Ordered manifest of logical export blocks — extend here when new dashboard widgets ship.
 */
export function buildCrcPptxWidgetPlan(snapshot: CrcPptxDashboardSnapshot): CrcPptxWidgetDescriptor[] {
  const d = snapshot;
  const plan: CrcPptxWidgetDescriptor[] = [];

  if (d.exportPptx.cover) {
    plan.push({ id: "cover", kind: "cover", title: "Couverture" });
  }

  if (d.exportPptx.kpi) {
    plan.push({
      id: "kpi-summary",
      kind: "kpi",
      title: "Synthèse KPI",
      kpiKeys: Object.keys(d.kpis).filter((k) => d.kpis[k as keyof typeof d.kpis]) as CrcPptxWidgetDescriptor["kpiKeys"],
    });
  }

  if (d.charts.geoBars || d.charts.geoDonut) {
    plan.push({
      id: "regional",
      kind: "chart",
      title: "Analytique régionale",
      chartKeys: ["geoBars", "geoDonut"].filter((k) => d.charts[k as "geoBars" | "geoDonut"]),
    });
  }

  if (d.charts.regionCards) {
    plan.push({
      id: "region-result-cards",
      kind: "chart",
      title: "Régions — résultats (cartes)",
      chartKeys: ["regionCards"],
    });
  }

  if (d.exportPptx.results) {
    plan.push({
      id: "resultat",
      kind: "composite",
      title: "Résultat & tendances",
      chartKeys: ["statusPie", "dailyArea", "trendLine", "monthlyBars"].filter((k) => d.charts[k as keyof typeof d.charts]),
      tableKeys: d.tables.pivotResult ? ["pivotResult"] : [],
    });
  }

  if (d.exportPptx.metier) {
    if (d.tables.pivotMetier) {
      plan.push({
        id: "metier",
        kind: "composite",
        title: "Métier × régions",
        tableKeys: ["pivotMetier"],
      });
    }
    if (d.tables.pivotNature) {
      plan.push({
        id: "nature",
        kind: "composite",
        title: "Nature de réclamation × régions",
        tableKeys: ["pivotNature"],
      });
    }
  }

  if (d.exportPptx.operators) {
    plan.push({
      id: "teleop",
      kind: "composite",
      title: "Téléopérateurs",
      chartKeys: d.charts.teleopBars ? ["teleopBars"] : [],
      tableKeys: d.tables.teleOpStats ? ["teleOpStats"] : [],
    });
  }

  if (d.tables.rawPreview) {
    plan.push({
      id: "raw-preview",
      kind: "table",
      title: "Grille brute (extrait)",
      rawColumnKeys: Object.keys(d.rawColumns).filter((k) => d.rawColumns[k as keyof typeof d.rawColumns]) as CrcPptxWidgetDescriptor["rawColumnKeys"],
    });
  }

  if (d.exportPptx.definitions) {
    plan.push({ id: "definitions", kind: "definitions", title: "Définitions & mensuel" });
  }

  return plan;
}
