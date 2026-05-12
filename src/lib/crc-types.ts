export type RawCell = string | number | boolean | Date | null | undefined;

import type { CanonicalRegion } from "./crc-constants";

/** Normalized CRC row — every row preserved, no deduplication */
export interface CrcRow {
  rawIndex: number;
  valid: boolean;
  validationReason?: string;

  campagneNom: string;
  campagneType: string;
  formulaire: string;
  date: Date | null;

  /**
   * Original Excel text value preserved exactly.
   * Example:
   * 10/05/2026 à 21:13
   */
  rawDateText: string;
  moisLabel: string;
  téléopérateur: string;
  /**
   * Libellé Résultat **canonique** (normalisé à l’import via `normalizeResult`).
   * Toutes agrégations / filtres / graphes utilisent ce champ.
   */
  résultat: string;
  /** Valeur Excel brute (colonne Résultat), inchangée pour audit / colonne « Excel ». */
  résultatRaw: string;
  tempsAttente: string;
  tempsAttenteIvr: string;
  tempsAttenteQueue: string;
  téléopérateurNote: string;
  téléphone: string;
  adresse: string;
  nomPrénom: string;
  nIdentité: string;
  police: string;
  natureRéclamation: string;
  page3Type: string;
  regions: string;
  provinces: string;
  communes: string;
  metier: string;

  /** Région for analytics — empty geographic → Faux Appels */
  régionCanon: CanonicalRegion;
}

export interface HeaderMappingEntry {
  original: string;
  normalizedKey: string;
  /** Label shown in cockpit (colonnes lisibles Axilus) */
  displayLabel: string;
}

export interface ParseDebug {
  detectedHeaders: string[];
  normalizedHeaders: HeaderMappingEntry[];
  totalRowsRaw: number;
  parsedRows: number;
  validRows: number;
  invalidRows: number;
  sheets: string[];
  logs: string[];
}

export interface ParseResult {
  rows: CrcRow[];
  debug: ParseDebug;
}
