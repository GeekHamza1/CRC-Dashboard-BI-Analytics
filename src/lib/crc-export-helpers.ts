import { REGION_ORDER, REGION_SHORT } from "./crc-constants";
import type { RawColumnKey } from "./crc-report-config";
import type { CrcRow } from "./crc-types";

export type OperatorRankRow = {
  name: string;
  volume: number;
  abandons: number;
  appelsDécrochésInterrompus: number;
  informés: number;
  tickets: number;
};

/** Re-export shape from analytics — avoid circular import by defining minimal type */
export type PivotRegionRow = {
  name?: string;
  métier?: string;
  nature?: string;
  [regionShort: string]: string | number | undefined;
};

const TELE_KEYS = ["volume", "appelsDécrochésInterrompus", "informés", "tickets"] as const;
export type TeleOpMetricKey = (typeof TELE_KEYS)[number];

export const TELEOP_COLUMN_DEFS: { key: TeleOpMetricKey | "name"; label: string }[] = [
  { key: "name", label: "Téléopérateur" },
  { key: "volume", label: "Volume" },
  { key: "appelsDécrochésInterrompus", label: "Appels décrochés interrompus" },
  { key: "informés", label: "Clients informés" },
  { key: "tickets", label: "Tickets transmis" },
];

export type RegionPivotId = "pivotResult" | "pivotMetier" | "pivotNature";

/** Snapshot passed into full-report exports (PDF / Excel / PPTX). */
export type CrcExportColumnVisibility = {
  pivotResultRegions?: Record<string, boolean>;
  pivotMetierRegions?: Record<string, boolean>;
  pivotNatureRegions?: Record<string, boolean>;
  teleOpMetrics?: Record<string, boolean>;
};

export type CrcColumnVisibilityState = {
  /** REGION_SHORT → visible (default true if absent) */
  pivotRegions: Record<RegionPivotId, Record<string, boolean>>;
  teleOpMetrics: Record<string, boolean>;
};

export function defaultRegionVisibility(): Record<string, boolean> {
  return Object.fromEntries(REGION_ORDER.map((r) => [REGION_SHORT[r], true]));
}

export function defaultTeleOpVisibility(): Record<string, boolean> {
  return Object.fromEntries(TELE_KEYS.map((k) => [k, true]));
}

export function defaultCrcColumnVisibility(): CrcColumnVisibilityState {
  const r = defaultRegionVisibility();
  return {
    pivotRegions: {
      pivotResult: { ...r },
      pivotMetier: { ...r },
      pivotNature: { ...r },
    },
    teleOpMetrics: defaultTeleOpVisibility(),
  };
}

export function activeRegionShorts(mask: Record<string, boolean> | undefined): string[] {
  const all = REGION_ORDER.map((rg) => REGION_SHORT[rg]);
  const t = all.filter((s) => mask?.[s] !== false);
  return t.length ? t : all;
}

export function activeTeleOpKeys(mask: Record<string, boolean> | undefined): TeleOpMetricKey[] {
  const t = TELE_KEYS.filter((k) => mask?.[k] !== false);
  return t.length ? [...t] : [...TELE_KEYS];
}

export function numericValue(obj: Record<string, unknown>, key: string): number {
  const value = obj[key];
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
}

/** Pivot row cells for visible regions + total of visible only */
export function pivotRowToCells(row: PivotRegionRow, shorts: string[]): { cells: string[]; total: number } {
  const nums = shorts.map((s) => numericValue(row, s));
  return {
    cells: nums.map(String),
    total: nums.reduce((a, b) => a + b, 0),
  };
}

export function filterOperatorRow(
  o: OperatorRankRow,
  metrics: TeleOpMetricKey[],
): Record<string, string | number> {
  const out: Record<string, string | number> = { name: o.name };
  for (const k of metrics) {
    if (k === "appelsDécrochésInterrompus") out[k] = o.appelsDécrochésInterrompus ?? 0;
    else if (k === "volume") out[k] = o.volume;
    else if (k === "informés") out[k] = o.informés;
    else if (k === "tickets") out[k] = o.tickets;
  }
  return out;
}

/** Valeurs texte pour exports tableaux (Excel / PDF / PPTX) — aligné sur la grille cockpit. */
export function formatRawCellForExport(r: CrcRow, key: RawColumnKey): string {
  switch (key) {
    case "date":
      return r.date ? r.date.toLocaleDateString("fr-FR") : "";
    case "resultatRaw":
      return r.résultatRaw;
    case "teleop":
      return r.téléopérateur;
    case "metier":
      return r.metier;
    case "nature":
      return r.natureRéclamation;
    case "regionCanon":
      return REGION_SHORT[r.régionCanon];
    case "regionsource":
      return r.regions;
    case "valid":
      return r.valid ? "oui" : "non";
    case "phone":
      return r.téléphone;
    default:
      return "";
  }
}
