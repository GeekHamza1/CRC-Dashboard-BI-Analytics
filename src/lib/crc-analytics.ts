import {
  CHART_PALETTE,
  NON_RENSEIGNE,
  REGION_COLORS,
  REGION_ORDER,
  REGION_SHORT,
  type CanonicalRegion,
} from "./crc-constants";
import type { CrcRow } from "./crc-types";

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

function nk(v: string) {
  return (v || "")
    .normalize("NFKD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .trim();
}

type StatusClass =
  | "abandon"
  | "appel_abandonne"
  | "client_informe"
  | "ticket_transmis"
  | "other";

function classifyStatus(label: string): StatusClass {
  const rawTrim = (label || "").trim();
  const s = nk(rawTrim);
  if (!s) return "other";

  /** Match Axilus libellés on **raw/normalized strings** — ordre garantit distinction Abandon ≠ Appel abandonné */
  const isTicket =
    s === nk("Ticket Transmis") || (s.includes("ticket") && (s.includes("transmis") || s.includes("transfer")));
  const isClientInformé =
    s === nk("Le client Informé") ||
    (s.includes("client") && (s.includes("inform") || s.includes("infor")));
  const isAbandonPur =
    s === nk("Abandon") || (s.endsWith("abandon") && !s.includes("appel") && !s.includes("appels"));
  const isAppelAbandonne =
    s === nk("Appel Abandonné") ||
    s === nk("Appels Abandonnés") ||
    (s.includes("appel") && s.includes("aband")) ||
    (s.includes("decroch") && (s.includes("interrompu") || s.includes("interromp") || s.includes("interrupt")));

  if (isTicket) return "ticket_transmis";
  if (isClientInformé) return "client_informe";

  /** Appel abandonné doit primer sur abandon large contenant « aband » mais sans « appel » */
  if (isAppelAbandonne && !isAbandonPur) return "appel_abandonne";
  if (isAbandonPur) return "abandon";

  if (s === nk("Abandon") || (!s.includes("appel") && !s.includes("appels") && s.includes("aband"))) {
    return "abandon";
  }
  return "other";
}

const STATUS_LABELS: Record<StatusClass, string> = {
  abandon: "Appels abandonnés",
  appel_abandonne: "Appels décrochés interrompus",
  client_informe: "Clients informés",
  ticket_transmis: "Tickets transmis",
  other: NON_RENSEIGNE,
};

/** Dynamic bucket with pattern-based normalization, no strict equals dependency. */
export function résultatBucket(label: string) {
  const raw = (label || "").trim();
  if (!raw) return NON_RENSEIGNE;
  const klass = classifyStatus(raw);
  if (klass === "other") return raw;
  return STATUS_LABELS[klass];
}

/** Keep every row identity; filter out only excluded by UI state */
export function applyFilters(rows: CrcRow[], f: DashboardFilters): CrcRow[] {
  return rows.filter((r) => {
    if (f.onlyValid && !r.valid) return false;
    if (!f.régions.includes(r.régionCanon)) return false;
    if (f.téléopérateurs.length && !f.téléopérateurs.includes(r.téléopérateur)) return false;
    if (f.résultats.length && !f.résultats.includes(r.résultat)) return false;
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

export function globalKpis(rows: CrcRow[]) {
  const abandons = rows.filter((r) => classifyStatus(r.résultat) === "abandon").length;
  const informés = rows.filter((r) => classifyStatus(r.résultat) === "client_informe").length;
  const tickets = rows.filter((r) => classifyStatus(r.résultat) === "ticket_transmis").length;
  return {
    totalRows: rows.length,
    réclamations: rows.length,
    totalAppels: rows.length,
    appelsAbandonnés: abandons,
    clientsInformés: informés,
    ticketsTransmis: tickets,
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
  const keys = [...new Set(rows.map((r) => résultatBucket(r.résultat)))];
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
    matrix[k] = Object.fromEntries(REGION_ORDER.map((r) => [REGION_SHORT[r], 0])) as Record<
      string,
      number
    >;
  });
  rows.forEach((r) => {
    const rk = résultatBucket(r.résultat);
    matrix[rk] ??= Object.fromEntries(REGION_ORDER.map((x) => [REGION_SHORT[x], 0])) as Record<
      string,
      number
    >;
    matrix[rk][REGION_SHORT[r.régionCanon]] += 1;
  });
  return keys.map((k) => ({ name: k, ...matrix[k] }));
}

export function pivotMétierParRégion(rows: CrcRow[]) {
  const metiers = [...new Set(rows.map((r) => r.metier))].sort((a, b) => a.localeCompare(b, "fr"));
  return metiers.map((m) => {
    const out: Record<string, number | string> = { métier: m };
    REGION_ORDER.forEach((rg) => {
      out[REGION_SHORT[rg]] = rows.filter((r) => r.metier === m && r.régionCanon === rg).length;
    });
    return out;
  });
}

export function pivotNatureParRégion(rows: CrcRow[]) {
  const natures = [...new Set(rows.map((r) => r.natureRéclamation))].sort((a, b) =>
    a.localeCompare(b, "fr"),
  );
  return natures.map((n) => {
    const out: Record<string, number | string> = { nature: n };
    REGION_ORDER.forEach((rg) => {
      out[REGION_SHORT[rg]] = rows.filter(
        (r) => r.natureRéclamation === n && r.régionCanon === rg,
      ).length;
    });
    return out;
  });
}

function zeroRegionCounts(): Record<CanonicalRegion, number> {
  return Object.fromEntries(REGION_ORDER.map((r) => [r, 0])) as Record<CanonicalRegion, number>;
}

export function dailySeries(rows: CrcRow[]) {
  const bucket = new Map<string, Record<CanonicalRegion, number>>();
  rows.forEach((r) => {
    if (!r.date) return;
    const key = ymd(r.date);
    const hit = bucket.get(key) ?? zeroRegionCounts();
    hit[r.régionCanon] += 1;
    bucket.set(key, hit);
  });
  return [...bucket.keys()].sort().map((jour) => {
    const d = bucket.get(jour)!;
    const out: Record<string, string | number> = { jour, total: 0 };
    REGION_ORDER.forEach((rg) => {
      out[REGION_SHORT[rg]] = d[rg];
      out.total = Number(out.total) + d[rg];
    });
    return out;
  });
}

export function monthlySeries(rows: CrcRow[]) {
  const bucket = new Map<string, Record<CanonicalRegion, number>>();
  rows.forEach((r) => {
    const month =
      r.moisLabel?.trim() && r.moisLabel !== NON_RENSEIGNE
        ? r.moisLabel
        : r.date
          ? `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`
          : "";
    if (!month) return;
    const hit = bucket.get(month) ?? zeroRegionCounts();
    hit[r.régionCanon] += 1;
    bucket.set(month, hit);
  });
  return [...bucket.keys()].sort().map((mois) => {
    const d = bucket.get(mois)!;
    const out: Record<string, string | number> = { mois, total: 0 };
    REGION_ORDER.forEach((rg) => {
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
    résultat: [...countMap(slice.map((r) => résultatBucket(r.résultat))).entries()].sort(
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
