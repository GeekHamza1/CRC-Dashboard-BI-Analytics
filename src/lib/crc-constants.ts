/** Canonical region keys used for consistent coloring */
export const REGION_ORDER = [
  "Drâa-Tafilalet",
  "Laâyoune-Sakia El Hamra",
  "Souss-Massa",
  "Inconnu",
] as const;

export type CanonicalRegion = (typeof REGION_ORDER)[number];

/** Hex colors per spec */
export const REGION_COLORS: Record<CanonicalRegion, string> = {
  "Drâa-Tafilalet": "#f97316",
  "Laâyoune-Sakia El Hamra": "#22c55e",
  "Souss-Massa": "#92400e",
  "Inconnu": "#94a3b8",
};

/** Short labels for compact chart axes */
export const REGION_SHORT: Record<CanonicalRegion, string> = {
  "Drâa-Tafilalet": "Drâa",
  "Laâyoune-Sakia El Hamra": "Laâyoune",
  "Souss-Massa": "Souss Massa",
  "Inconnu": "Inconnu",
};

export const KEY_OPERATIONAL_STATUSES = [
  "Abandon",
  "Appel Abandonné",
  "Le client Informé",
  "Ticket Transmis",
] as const;

export const NON_RENSEIGNE = "Non renseigné";
export const FAUX_APPELS = "Inconnu";

/** Professional multi-hue palette for series (light & dark safe) */
export const CHART_PALETTE = [
  "#3b82f6",
  "#06b6d4",
  "#10b981",
  "#f97316",
  "#ef4444",
  "#eab308",
  "#a855f7",
  "#ec4899",
  "#14b8a6",
  "#6366f1",
];

export function regionToCanonical(raw: string): CanonicalRegion {
  const t = raw.trim();
  if (!t || t === FAUX_APPELS) return "Inconnu";
  const lower = t.toLowerCase();
  if (lower.includes("drâa") || lower.includes("draa") || lower.includes("tafilalet"))
    return "Drâa-Tafilalet";
  if (
    lower.includes("laâyoune") ||
    lower.includes("laayoune") ||
    lower.includes("sakia") ||
    lower.includes("hamra")
  )
    return "Laâyoune-Sakia El Hamra";
  if (lower.includes("souss") || lower.includes("massa")) return "Souss-Massa";
  return "Inconnu";
}
