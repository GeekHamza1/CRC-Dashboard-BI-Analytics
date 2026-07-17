import {
  CHART_PALETTE,
  NON_RENSEIGNE,
  REGION_COLORS,
  REGION_ORDER,
  REGION_SHORT,
  type CanonicalRegion,
} from "./crc-constants";
import { normalizeResult, normalizeResultKey } from "./crc-normalize-result";
import { compareResultBuckets } from "./constants/chart-colors";
import type { CrcRow } from "./crc-types";

export { CANONICAL_RESULT_LABELS, normalizeResult } from "./crc-normalize-result";

export interface DashboardFilters {
  onlyValid: boolean;
  régions: CanonicalRegion[];
  téléopérateurs: string[];
  résultats: string[];
  dateFrom: string | null;
  dateTo: string | null;
}

export const defaultDashboardFilters = (): DashboardFilters => ({
  onlyValid: false,
  régions: [...REGION_ORDER],
  téléopérateurs: [],
  résultats: [],
  dateFrom: null,
  dateTo: null,
});

function ymd(ts: Date) {
  return `${ts.getFullYear()}-${String(ts.getMonth() + 1).padStart(2, "0")}-${String(ts.getDate()).padStart(2, "0")}`;
}

/** Champs Excel vides / placeholders (ne pas compter comme renseignés). */
export function isEmptyField(v: unknown): boolean {
  if (v == null) return true;
  const s = normalizeResultKey(String(v));
  if (s === "") return true;
  if (s === "-") return true;
  if (s === normalizeResultKey(NON_RENSEIGNE)) return true;
  return false;
}

type StatusClass =
  | "abandon"
  | "appel_abandonne"
  | "client_informe"
  | "ticket_transmis"
  | "other";

/** Famille KPI / agrégations — dérivée de `normalizeResult` uniquement. */
export function classifyStatus(label: string): StatusClass {
  const canon = normalizeResult(label ?? "");
  if (canon === "Appels abandonnés") return "abandon";
  if (canon === "Appels décrochés interrompus") return "appel_abandonne";
  if (canon === "Clients informés") return "client_informe";
  if (canon === "Tickets transmis") return "ticket_transmis";
  return "other";
}

/** Libellé bucket agrégé (synonymes Axilus + libellés dashboard → canoniques). */
export function résultatBucket(label: string): string {
  const raw = String(label ?? "").trim();
  if (!raw) return NON_RENSEIGNE;
  const n = normalizeResult(raw);
  return n || NON_RENSEIGNE;
}

/** Keep every row identity; filter out only excluded by UI state */
export function applyFilters(rows: CrcRow[], f: DashboardFilters): CrcRow[] {
  return rows.filter((r) => {
    if (f.onlyValid && !r.valid) return false;
    if (!f.régions.includes(r.régionCanon)) return false;
    if (f.téléopérateurs.length && !f.téléopérateurs.includes(r.téléopérateur)) return false;
    if (f.résultats.length) {
      const rowCanon = r.résultat;
      const keep = f.résultats.some((sel) => normalizeResult(sel) === rowCanon);
      if (!keep) return false;
    }
    if (r.date) {
      const d = ymd(r.date);
      if (f.dateFrom && d < f.dateFrom) return false;
      if (f.dateTo && d > f.dateTo) return false;
    } else if (f.dateFrom || f.dateTo) return false;
    return true;
  });
}

export function countMap(keys: string[]) {
  const m = new Map<string, number>();
  for (const k of keys) {
    const kk = (k || "").trim() || NON_RENSEIGNE;
    m.set(kk, (m.get(kk) ?? 0) + 1);
  }
  return m;
}

export function activeCanonicalRegions(rows: CrcRow[]): CanonicalRegion[] {
  return REGION_ORDER.filter((region) => rows.some((row) => row.régionCanon === region));
}

export function parseQueueSeconds(value: string | null | undefined): number | null {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;

  const normalized = raw.replace(/\s+/g, "");
  if (/^\d+$/.test(normalized)) return Number(normalized);

  const parts = normalized.split(":").map((part) => Number(part));
  if (parts.some((part) => Number.isNaN(part))) return null;

  if (parts.length === 1) return parts[0];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  return null;
}

export function rowHasQueueWait(row: CrcRow): boolean {
  const secs = parseQueueSeconds(row.tempsAttenteQueue);
  return secs != null && secs > 1;
}

export function formatDurationSeconds(totalSeconds: number): string {
  if (!Number.isFinite(totalSeconds) || totalSeconds < 0) return "—";
  const safeSeconds = Math.max(0, Math.floor(totalSeconds));
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const seconds = safeSeconds % 60;
  const hh = String(hours).padStart(2, "0");
  const mm = String(minutes).padStart(2, "0");
  const ss = String(seconds).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

export interface ShiftResultBucket extends Record<string, number | string> {
  label: string;
  count: number;
}

export interface ShiftResultDistribution {
  data: ShiftResultBucket[];
  resultLabels: string[];
}

export function shiftBuckets(rows: CrcRow[]) {
  const buckets = [
    { label: "00h–08h", min: 0, max: 7, count: 0 },
    { label: "08h–16h", min: 8, max: 15, count: 0 },
    { label: "16h–00h", min: 16, max: 23, count: 0 },
  ];

  rows.forEach((row) => {
    if (!row.date) return;
    const hour = row.date.getHours();
    const bucket = buckets.find((b) => hour >= b.min && hour <= b.max);
    if (bucket) bucket.count += 1;
  });

  return buckets;
}

export function shiftResultDistribution(rows: CrcRow[]): ShiftResultDistribution {
  const buckets = [
    { label: "00h–08h", min: 0, max: 7 },
    { label: "08h–16h", min: 8, max: 15 },
    { label: "16h–00h", min: 16, max: 23 },
  ];

  const resultLabels = new Set<string>();
  rows.forEach((row) => {
    if (!row.date) return;
    const result = row.résultat || NON_RENSEIGNE;
    resultLabels.add(result);
  });

  const sortedResultLabels = [...resultLabels].sort(compareResultBuckets);
  const distribution: ShiftResultBucket[] = buckets.map((bucket) => {
    const row: ShiftResultBucket = {
      label: bucket.label,
      count: 0,
    };
    sortedResultLabels.forEach((result) => {
      row[result] = 0;
    });
    return row;
  });

  rows.forEach((row) => {
    if (!row.date) return;
    const hour = row.date.getHours();
    const bucket = distribution.find((b, index) => {
      const def = buckets[index];
      return hour >= def.min && hour <= def.max;
    });
    if (!bucket) return;

    const result = row.résultat || NON_RENSEIGNE;
    bucket.count += 1;
    bucket[result] = (bucket[result] as number) + 1;
  });

  return {
    data: distribution,
    resultLabels: sortedResultLabels,
  };
}

export function hourlyCallDistribution(rows: CrcRow[]) {
  const counts = new Array(24).fill(0);
  rows.forEach((row) => {
    if (!row.date) return;
    counts[row.date.getHours()] += 1;
  });
  return counts.map((count, hour) => ({
    hour: `${String(hour).padStart(2, "0")}h`,
    count,
  }));
}

export function globalKpis(rows: CrcRow[]) {
  const abandons = rows.filter((r) => classifyStatus(r.résultat) === "abandon").length;
  const appelsDecrochesInterrompus = rows.filter(
    (r) => classifyStatus(r.résultat) === "appel_abandonne",
  ).length;
  const informés = rows.filter((r) => classifyStatus(r.résultat) === "client_informe").length;
  const tickets = rows.filter((r) => classifyStatus(r.résultat) === "ticket_transmis").length;
  const queueSeconds = rows
    .map((r) => parseQueueSeconds(r.tempsAttenteQueue))
    .filter((value): value is number => value != null && value > 1);
  const totalClientsWaited = queueSeconds.length;
  const avgWaitingSeconds = totalClientsWaited
    ? queueSeconds.reduce((sum, value) => sum + value, 0) / totalClientsWaited
    : 0;
  const pctClientsWaited = rows.length ? (totalClientsWaited / rows.length) * 100 : 0;
  return {
    totalRows: rows.length,
    réclamations: rows.length,
    totalAppels: rows.length,
    appelsAbandonnés: abandons,
    appelsDécrochésInterrompus: appelsDecrochesInterrompus,
    clientsInformés: informés,
    ticketsTransmis: tickets,
    avgWaitingTimeSeconds: avgWaitingSeconds,
    avgWaitingTime: formatDurationSeconds(avgWaitingSeconds),
    totalClientsWaited,
    pctClientsWaited,
    appelsParRégion: countMap(rows.map((r) => r.régionCanon)),
    appelsParMétier: countMap(rows.map((r) => r.metier)),
    appelsParTéléop: countMap(rows.map((r) => r.téléopérateur)),
  };
}

export function pivotRésultatParRégion(rows: CrcRow[]) {
  const order = [
    "Appels abandonnés",
    "Appels décrochés interrompus",
    "Clients informés",
    "Tickets transmis",
  ];
  const regions = activeCanonicalRegions(rows);
  const keys = [...new Set(rows.map((r) => r.résultat))];
  keys.sort((a, b) => {
    const ia = order.indexOf(a);
    const ib = order.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return a.localeCompare(b, "fr");
  });
  const matrix: Record<string, Record<string, number>> = {};
  keys.forEach((k) => {
    matrix[k] = Object.fromEntries(regions.map((r) => [REGION_SHORT[r], 0])) as Record<
      string,
      number
    >;
  });
  rows.forEach((r) => {
    const rk = r.résultat;
    matrix[rk] ??= Object.fromEntries(regions.map((x) => [REGION_SHORT[x], 0])) as Record<
      string,
      number
    >;
    matrix[rk][REGION_SHORT[r.régionCanon]] += 1;
  });
  return keys.map((k) => ({ name: k, ...matrix[k] }));
}

export function pivotMétierParRégion(rows: CrcRow[]) {
  const metiers = [...new Set(rows.map((r) => r.metier))].sort((a, b) => a.localeCompare(b, "fr"));
  const regions = activeCanonicalRegions(rows);
  return metiers.map((m) => {
    const out: Record<string, number | string> = { métier: m };
    regions.forEach((rg) => {
      out[REGION_SHORT[rg]] = rows.filter((r) => r.metier === m && r.régionCanon === rg).length;
    });
    return out;
  });
}

export function pivotNatureParRégion(rows: CrcRow[]) {
  const natures = [...new Set(rows.map((r) => r.natureRéclamation))].sort((a, b) =>
    a.localeCompare(b, "fr"),
  );
  const regions = activeCanonicalRegions(rows);
  return natures.map((n) => {
    const out: Record<string, number | string> = { nature: n };
    regions.forEach((rg) => {
      out[REGION_SHORT[rg]] = rows.filter(
        (r) => r.natureRéclamation === n && r.régionCanon === rg,
      ).length;
    });
    return out;
  });
}

/** Provinces aggregation per region — returns map of region → sorted provinces list */
export function provincesParRégion(rows: CrcRow[]) {
  const regions = activeCanonicalRegions(rows);
  const result: Record<CanonicalRegion, [string, number][]> = Object.fromEntries(
    regions.map((region) => [region, [] as [string, number][]]),
  ) as Record<CanonicalRegion, [string, number][]>;
  
  regions.forEach((region) => {
    const regionRows = rows.filter((r) => r.régionCanon === region);
    const provinces = countMap(regionRows.map((r) => r.provinces));
    result[region] = [...provinces.entries()].sort((a, b) => b[1] - a[1]);
  });
  
  return result;
}

function zeroRegionCounts(rows: CrcRow[]): Record<CanonicalRegion, number> {
  return Object.fromEntries(activeCanonicalRegions(rows).map((r) => [r, 0])) as Record<CanonicalRegion, number>;
}

export function dailySeries(rows: CrcRow[]) {
  const bucket = new Map<string, Record<CanonicalRegion, number>>();
  const regions = activeCanonicalRegions(rows);
  rows.forEach((r) => {
    if (!r.date) return;
    const key = ymd(r.date);
    const hit = bucket.get(key) ?? zeroRegionCounts(rows);
    hit[r.régionCanon] += 1;
    bucket.set(key, hit);
  });
  return [...bucket.keys()].sort().map((jour) => {
    const d = bucket.get(jour)!;
    const out: Record<string, string | number> = { jour, total: 0 };
    regions.forEach((rg) => {
      out[REGION_SHORT[rg]] = d[rg];
      out.total = Number(out.total) + d[rg];
    });
    return out;
  });
}
const MONTH_INDEX_FR: Record<string, number> = {
  janvier: 0,
  février: 1,
  fevrier: 1,
  mars: 2,
  avril: 3,
  mai: 4,
  juin: 5,
  juillet: 6,
  août: 7,
  aout: 7,
  septembre: 8,
  octobre: 9,
  novembre: 10,
  décembre: 11,
  decembre: 11,
};

function parseMonthLabel(label: string): number {
  const value = String(label ?? "").trim();

  // FORMAT: 2026-04
  const isoMatch = value.match(/^(\d{4})-(\d{2})$/);

  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]) - 1;
    return new Date(year, month, 1).getTime();
  }

  // FORMAT: Avril 2026
  const parts = value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .split(/\s+/);

  if (parts.length >= 2) {
    const monthName = parts[0];
    const year = Number(parts[1]);

    const monthIndex = MONTH_INDEX_FR[monthName];

    if (
      typeof monthIndex === "number" &&
      Number.isFinite(year)
    ) {
      return new Date(year, monthIndex, 1).getTime();
    }
  }

  return 0;
}

export function monthlySeries(rows: CrcRow[]) {
  const bucket = new Map<string, Record<CanonicalRegion, number>>();
  const regions = activeCanonicalRegions(rows);
  rows.forEach((r) => {
    const month =
      r.moisLabel?.trim() && r.moisLabel !== NON_RENSEIGNE
        ? r.moisLabel
        : r.date
          ? `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`
          : "";
          if (
        !month ||
        !String(month).trim() ||
        parseMonthLabel(String(month)) === 0
      ) {
        return;
      }
    const hit = bucket.get(month) ?? zeroRegionCounts(rows);
    hit[r.régionCanon] += 1;
    bucket.set(month, hit);
  });
return [...bucket.keys()]
  .sort((a, b) => parseMonthLabel(a) - parseMonthLabel(b))
  .map((mois) => {
    const d = bucket.get(mois)!;
    const out: Record<string, string | number> = { mois, total: 0 };
    regions.forEach((rg) => {
      out[REGION_SHORT[rg]] = d[rg];
      out.total = Number(out.total) + d[rg];
    });
    return out;
  });
}

export function operatorRanking(rows: CrcRow[]) {
  const byOperator = countMap(rows.map((r) => r.téléopérateur));
  return [...byOperator.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([name, volume]) => {
      const scoped = rows.filter((r) => r.téléopérateur === name);
      const abandons = scoped.filter((r) => classifyStatus(r.résultat) === "abandon").length;
      const appelsDécrochésInterrompus = scoped.filter(
        (r) => classifyStatus(r.résultat) === "appel_abandonne",
      ).length;
      const informés = scoped.filter((r) => classifyStatus(r.résultat) === "client_informe").length;
      const tickets = scoped.filter((r) => classifyStatus(r.résultat) === "ticket_transmis").length;
      return {
        name,
        volume,
        abandons,
        appelsDécrochésInterrompus,
        informés,
        tickets,
      };
    });
}

/** Region slice analytics */
export function regionSlice(rows: CrcRow[], region: CanonicalRegion) {
  const slice = rows.filter((r) => r.régionCanon === region);
  return {
    total: slice.length,
    résultat: [...countMap(slice.map((r) => r.résultat)).entries()].sort(
      (a, b) => b[1] - a[1],
    ),
    métier: [...countMap(slice.map((r) => r.metier)).entries()].sort((a, b) => b[1] - a[1]),
    nature: [...countMap(slice.map((r) => r.natureRéclamation)).entries()].sort(
      (a, b) => b[1] - a[1],
    ),
    daily: dailySeries(slice),
    monthly: monthlySeries(slice),
  };
}

/** Theme-aware palettes for SVG */
export function chartSemanticPalette(isDark: boolean) {
  return {
    fg: isDark ? "#f8fafc" : "#0f172a",
    muted: isDark ? "#94a3b8" : "#64748b",
    grid: isDark ? "rgba(248,250,252,0.06)" : "rgba(15,23,42,0.08)",
    card: isDark ? "rgba(255,255,255,0.04)" : "rgba(255,255,255,0.82)",
    series: [...CHART_PALETTE],
    region: REGION_ORDER.map((r) => REGION_COLORS[r]),
    tooltipBg: isDark ? "#020617ee" : "#ffffffee",
    legend: REGION_SHORT,
  };
}

export { REGION_COLORS, REGION_ORDER, REGION_SHORT };
