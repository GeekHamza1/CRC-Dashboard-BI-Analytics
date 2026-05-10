/**
 * Palette globale Résultat — couleurs fixes et dynamiques.
 * Compatible avec Axilus CRC dynamique.
 */

import { normalizeResult } from "../crc-normalize-result";
import { NON_RENSEIGNE } from "../crc-constants";

/**
 * Couleurs OFFICIELLES des résultats connus.
 * Les autres valeurs dynamiques utilisent RESULT_OTHER_FALLBACK.
 */
export const RESULT_COLORS = {
  "Appels abandonnés": "#EF4444",
  "Appels décrochés interrompus": "#F59E0B",
  "Clients informés": "#10B981",
  "Tickets transmis": "#3B82F6",
  [NON_RENSEIGNE]: "#6B7280",
} as const;

/**
 * Ordre préféré d'affichage dans tableaux/charts.
 * Les autres valeurs Axilus suivent automatiquement après.
 */
export const RESULT_DISPLAY_ORDER: readonly string[] = [
  "Appels abandonnés",
  "Appels décrochés interrompus",
  "Clients informés",
  "Tickets transmis",
  NON_RENSEIGNE,
];

/**
 * Couleur fallback pour nouveaux statuts Axilus inconnus.
 */
export const RESULT_OTHER_FALLBACK = "#9333EA";

/**
 * Volume global neutre.
 */
export const AGGREGATE_VOLUME_BAR = "#475569";

/**
 * KPI global abandons agrégés.
 */
export const OPERATOR_ABANDON_METRIC = "#EA580C";

export type KnownResultLabel = keyof typeof RESULT_COLORS;

/**
 * HEX → RGB tuple (PDF exports)
 */
export function hexToRgbTuple(
  hex: string,
): [number, number, number] {
  let h = hex.trim().replace("#", "");

  if (h.length === 3) {
    h = h
      .split("")
      .map((c) => c + c)
      .join("");
  }

  const n = Number.parseInt(h, 16);

  return [
    (n >> 16) & 255,
    (n >> 8) & 255,
    n & 255,
  ];
}

/**
 * Tri logique des résultats.
 */
export function compareResultBuckets(
  a: string,
  b: string,
): number {
  const order = RESULT_DISPLAY_ORDER;

  const ia = order.indexOf(a);
  const ib = order.indexOf(b);

  if (ia !== -1 && ib !== -1)
    return ia - ib;

  if (ia !== -1)
    return -1;

  if (ib !== -1)
    return 1;

  return a.localeCompare(b, "fr");
}

/**
 * Couleur dynamique pour tout résultat.
 * Fonctionne même avec nouveaux statuts Axilus.
 */
export function getResultColor(
  label: string,
): string {
  const trimmed = String(label ?? "").trim();
  const normalized = trimmed ? normalizeResult(trimmed) : NON_RENSEIGNE;

  if (
    normalized &&
    normalized in RESULT_COLORS
  ) {
    return RESULT_COLORS[
      normalized as KnownResultLabel
    ];
  }

  return RESULT_OTHER_FALLBACK;
}

/**
 * HEX compatible pptxgenjs.
 */
export function hexForPptx(hex: string) {
  return hex.replace(/^#/, "");
}