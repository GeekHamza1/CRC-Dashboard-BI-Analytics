/**
 * Normalisation « Résultat » — seule façade pour les libellés (parse + analytics).
 * Comparaisons par clé stable uniquement (pas de includes / fuzzy).
 */

/** Clé stable : accents, casse, espaces, NBSP. */
export function normalizeResultKey(v: string): string {
  return (v || "")
    .replace(/\u00a0/g, " ")
    .normalize("NFKD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

export const CANONICAL_RESULT_LABELS = [
  "Appels abandonnés",
  "Appels décrochés interrompus",
  "Clients informés",
  "Tickets transmis",
] as const;

/**
 * Synonymes Axilus + libellés dashboard déjà canoniques → libellé unique.
 * Valeurs inconnues : chaîne trimée inchangée.
 */
export function normalizeResult(raw: unknown): string {
  const trimmed = String(raw ?? "").trim();
  if (!trimmed) return "";

  const value = normalizeResultKey(trimmed);

  if (
    value === normalizeResultKey("Appel Abandonné") ||
    value === normalizeResultKey("Appel Abondonné") ||
    value === normalizeResultKey("Appels décrochés interrompus")
  ) {
    return "Appels décrochés interrompus";
  }
  if (value === normalizeResultKey("Abandon") || value === normalizeResultKey("Appels abandonnés")) {
    return "Appels abandonnés";
  }
  if (value === normalizeResultKey("Le client Informé") || value === normalizeResultKey("Clients informés")) {
    return "Clients informés";
  }
  if (value === normalizeResultKey("Ticket Transmis") || value === normalizeResultKey("Tickets transmis")) {
    return "Tickets transmis";
  }

  return trimmed;
}
