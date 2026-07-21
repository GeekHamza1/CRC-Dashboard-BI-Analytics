import * as XLSX from "xlsx";
import type { CrcExportColumnVisibility } from "./crc-export-helpers";
import { activeRegionShorts, activeTeleOpKeys, numericValue } from "./crc-export-helpers";
import type { CrcReportConfig } from "./crc-report-config";
import { mergeExportExcelSheets } from "./crc-report-config";
import {
  dailySeries,
  globalKpis,
  hourlyCallDistribution,
  monthlySeries,
  operatorRanking,
  pivotMétierParRégion,
  pivotNatureParRégion,
  pivotRésultatParRégion,
  rowHasQueueWait,
  shiftResultDistribution,
} from "./crc-analytics";
import { RESULT_COLORS } from "./constants/chart-colors";
import { REGION_ORDER, REGION_SHORT } from "./crc-constants";
import type { CrcRow } from "./crc-types";
import { reportTimestamp } from "./export-branding";

export function exportCrcExcel(
  rows: CrcRow[],
  basename = "crc_axilus_export",
  sheetPlan?: Partial<CrcReportConfig["exportExcelSheets"]>,
  columnVisibility?: CrcExportColumnVisibility,
) {
  const wb = XLSX.utils.book_new();
  const kpis = globalKpis(rows);
  const sh = mergeExportExcelSheets(sheetPlan);
  const visR = activeRegionShorts(columnVisibility?.pivotResultRegions);
  const visM = activeRegionShorts(columnVisibility?.pivotMetierRegions);
  const visN = activeRegionShorts(columnVisibility?.pivotNatureRegions);
  const teleK = activeTeleOpKeys(columnVisibility?.teleOpMetrics);

  if (sh.readme) {
    const intro = XLSX.utils.aoa_to_sheet([
      ["SRM - CRC Reporting Export"],
      ["Généré le", reportTimestamp()],
      ["Total Appels", rows.length],
      ["Style", "Operational CRC management workbook"],
      [],
      ["Code couleur Résultat", "Hex"],
      ...Object.entries(RESULT_COLORS),
    ]);
    XLSX.utils.book_append_sheet(wb, intro, "README");
  }

  if (sh.kpis) {
    const kpiAoA: (string | number)[][] = [
      ["Indicateur", "Valeur"],
      ["Total réclamations", kpis.réclamations],
      ["Total appels", kpis.totalAppels],
      ["Appels abandonnés", kpis.appelsAbandonnés],
      ["Clients informés", kpis.clientsInformés],
      ["Tickets transmis", kpis.ticketsTransmis],
      ["Temps d'attente moyen", kpis.avgWaitingTime],
      ["Appels orientés vers la file d’attente.", kpis.totalClientsWaited],
      ["% d'appels orientés vers la file d’attente.", `${kpis.pctClientsWaited.toFixed(1)} %`],
      ...REGION_ORDER.map((rg): [string, number] => [
        `Appels — ${REGION_SHORT[rg]}`,
        kpis.appelsParRégion.get(rg) ?? 0,
      ]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(kpiAoA), "Synthèse KPI");
  }

  if (sh.flat) {
    const flatData = rows.map((r) => ({
      CampagneNom: r.campagneNom,
      CampagneType: r.campagneType,
      Formulaire: r.formulaire,
      Date_ISO: r.date ? r.date.toISOString() : "",
      Mois: r.moisLabel,
      Téléopérateur: r.téléopérateur,
      Résultat: r.résultat,
      TempsAttente: r.tempsAttente,
      TempsAttenteIvr: r.tempsAttenteIvr,
      TempsAttenteQueue: r.tempsAttenteQueue,
      Téléphone: r.téléphone,
      Adresse: r.adresse,
      NomEtPrénom: r.nomPrénom,
      NIdentité: r.nIdentité,
      Police: r.police,
      NatureRéclamation: r.natureRéclamation,
      TypeComplaint: r.page3Type,
      RegionsSource: r.regions,
      Provinces: r.provinces,
      Communes: r.communes,
      Métier: r.metier,
      RégionCanon: r.régionCanon,
      LigneValide: r.valid ? "oui" : "non",
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(flatData), "Données brutes");
  }

  if (sh.pivotResult) {
    const pr = pivotRésultatParRégion(rows);
    const slim = pr.map((row) => {
      const o: Record<string, string | number> = { Résultat: row.name };
      for (const s of visR) o[s] = numericValue(row, s);
      o.Total = visR.reduce((acc, s) => acc + numericValue(row, s), 0);
      return o;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(slim), "Résultat par région");
  }

  if (sh.pivotMetier) {
    const pm = pivotMétierParRégion(rows);
    const slim = pm.map((row) => {
      const o: Record<string, string | number> = { Métier: row.métier };
      for (const s of visM) o[s] = numericValue(row, s);
      return o;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(slim), "Métier par région");
  }

  if (sh.pivotNature) {
    const pn = pivotNatureParRégion(rows);
    const slim = pn.map((row) => {
      const o: Record<string, string | number> = { Nature: row.nature };
      for (const s of visN) o[s] = numericValue(row, s);
      return o;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(slim), "Nature par région");
  }

  if (sh.operators) {
    const ops = operatorRanking(rows);
    const slim = ops.map((o) => {
      const r: Record<string, string | number> = { Téléopérateur: o.name };
      for (const k of teleK) r[k] = numericValue(o, k);
      return r;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(slim), "Classement des téléopérateurs");
  }

  if (sh.daily) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dailySeries(rows)), "Évolution journalière");
  }

  if (sh.hourly) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(hourlyCallDistribution(rows)),
      "Appels par heure",
    );
  }

  if (sh.shifts) {
    const shiftDist = shiftResultDistribution(rows);
    const shiftSheet = shiftDist.data.map((row) => {
      const out: Record<string, string | number> = { Shift: row.label, Total: row.count };
      shiftDist.resultLabels.forEach((label) => {
        out[label] = row[label] ?? 0;
      });
      return out;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(shiftSheet), "Shifts horaires");
  }

  if (sh.monthly) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlySeries(rows)), "Évolution mensuelle");
  }

  if (!wb.SheetNames.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Aucune feuille sélectionnée"]]), "Vide");
  }

  const fn = `${basename.replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`;
  XLSX.writeFile(wb, fn);
}
