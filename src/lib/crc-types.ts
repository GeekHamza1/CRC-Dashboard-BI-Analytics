export type RawCell = string | number | boolean | null | undefined;

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
  moisLabel: string;
  téléopérateur: string;
  résultat: string;
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
