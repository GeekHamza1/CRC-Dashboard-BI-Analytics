import * as XLSX from "xlsx";
import { normalizeResult } from "./crc-normalize-result";
import {
  FAUX_APPELS,
  NON_RENSEIGNE,
  regionToCanonical,
} from "./crc-constants";
import type {
  CrcRow,
  HeaderMappingEntry,
  ParseDebug,
  ParseResult,
  RawCell,
} from "./crc-types";

/** Internal field identifiers */
type RowField =
  | "campagneNom"
  | "campagneType"
  | "formulaire"
  | "date"
  | "mois"
  | "téléopérateur"
  | "résultat"
  | "tempsAttente"
  | "tempsAttenteIvr"
  | "tempsAttenteQueue"
  | "téléopérateurNote"
  | "téléphone"
  | "adresse"
  | "nomPrénom"
  | "nIdentité"
  | "police"
  | "natureRéclamation"
  | "page3Type"
  | "regions"
  | "provinces"
  | "communes"
  | "metier";

const FIELD_SYNONYMS: { field: RowField; match: RegExp[] }[] = [
  { field: "campagneNom", match: [/^campagne\/?nom$/i, /^nom\s*campagne$/i] },
  { field: "campagneType", match: [/^campagne\/?type$/i] },
  { field: "formulaire", match: [/^formulaire$/i] },
  { field: "date", match: [/^date$/i, /^.*\bdate\b.*réclamation$/i] },
  { field: "mois", match: [/^mois$/i] },
  {
    field: "téléopérateur",
    match: [
      /^téléop[eé]rateur$/i,
      /^teleop[eé]rateur$/i,
      /^teleoperator$/i,
      /^teleoperateur$/i,
    ],
  },
  { field: "résultat", match: [/^résultat$/i, /^resultat$/i] },
  { field: "tempsAttente", match: [/^temps d'attente$/i, /^temps\s*d['']attente$/i] },
  {
    field: "tempsAttenteIvr",
    match: [/^temps d'attente \(ivr\)$/i, /^temps\s*d['']attente\s*\(?ivr\)?$/i],
  },
  {
    field: "tempsAttenteQueue",
    match: [
      /^temps d'attente \(queue\)$/i,
      /^temps\s*d['']attente\s*\(?queue\)?$/i,
    ],
  },
  {
    field: "téléopérateurNote",
    match: [
      /^téléopérateur\/?note$/i,
      /^note\s*téléopérateur$/i,
      /^note\s*télé\s*$/i,
    ],
  },
  {
    field: "téléphone",
    match: [/^header\/?téléphone$/i, /^téléphone$/i],
  },
  { field: "adresse", match: [/^header\/?adresse$/i, /^adresse$/i] },
  {
    field: "nomPrénom",
    match: [/^header\/?nom\s*&\s*prénom$/i, /^nom\s*&\s*prénom$/i],
  },
  {
    field: "nIdentité",
    match: [
      /^header\/?n°\s*identité$/i,
      /^header\/?n[o°]\s*identit[eé]/i,
      /^n°\s*identité$/i,
    ],
  },
  {
    field: "police",
    match: [/^header\/?police$/i, /^police$/i],
  },
  {
    field: "natureRéclamation",
    match: [/^page3\/?nature de réclamation$/i, /^nature de réclamation$/i],
  },
  {
    field: "page3Type",
    match: [/^page3\/?type$/i, /^type$/i],
  },
  {
    field: "regions",
    match: [/^page3\/?regions?$/i, /^regions$/i, /^régions$/i],
  },
  {
    field: "provinces",
    match: [/^page3\/?provinces?$/i, /^provinces$/i],
  },
  {
    field: "communes",
    match: [/^page3\/?communes$/i, /^communes$/i],
  },
  {
    field: "metier",
    match: [
      /^page3\/?métier$/i,
      /^page3\/?metier$/i,
      /^métier$/i,
      /^metier$/i,
    ],
  },
];

export function displayLabel(header: string): string {
  let s = header.trim();
  if (s.startsWith("Page3/")) s = s.slice(6);
  if (s.startsWith("Header/")) s = s.slice(7);
  s = s.trim();
  const mapTitling: Record<string, string> = {
    Regions: "Régions",
    provinces: "Provinces",
    communes: "Communes",
  };
  const low = s.toLowerCase();
  if (mapTitling[s]) return mapTitling[s];
  if (low === "regions") return "Régions";
  return s;
}

function nk(s: string): string {
  return s
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function matchField(original: string): RowField | null {
  const key = nk(original);
  for (const { field, match } of FIELD_SYNONYMS) {
    for (const re of match) {
      if (re.test(original) || re.test(key)) return field;
      const nkUnderscore = key.replace(/\//g, " ");
      if (re.test(nkUnderscore)) return field;
    }
  }
  /** slash normalized */
  const slashFlat = key.replace(/\//g, " ");
  for (const { field, match } of FIELD_SYNONYMS) {
    for (const re of match) {
      if (re.test(slashFlat)) return field;
    }
  }
  return null;
}

function buildHeaderMap(headers: string[]): Map<number, RowField | null> {
  const map = new Map<number, RowField | null>();
  headers.forEach((h, i) => {
    map.set(i, matchField(h));
  });
  return map;
}

function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial)) return null;
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  if (Number.isNaN(dateInfo.getTime())) return null;

  const fractionalDay = serial - Math.floor(serial);
  const totalSeconds = Math.round(86400 * fractionalDay);
  dateInfo.setSeconds(dateInfo.getSeconds() + totalSeconds);
  return dateInfo;
}

function parseFlexibleDate(cell: RawCell): Date | null {
  if (cell == null || cell === "") {
    return null;
  }

  // Already a valid Date object
  if (
    typeof cell === "object" &&
    cell instanceof Date &&
    !Number.isNaN(cell.getTime())
  ) {
    return new Date(cell.getTime());
  }

  // Excel serial numbers
  if (typeof cell === "number" && Number.isFinite(cell)) {
    if (cell > 20000 && cell < 90000) {
      const parsed = XLSX.SSF.parse_date_code(cell);

      if (parsed) {
        return new Date(
          parsed.y,
          parsed.m - 1,
          parsed.d,
          parsed.H || 0,
          parsed.M || 0,
          parsed.S || 0
        );
      }
    }

    const timestampDate = new Date(cell);

    if (!Number.isNaN(timestampDate.getTime())) {
      return timestampDate;
    }
  }

  if (typeof cell === "boolean") {
    return null;
  }

  const text = String(cell).trim();

  if (!text) {
    return null;
  }

  /**
   * French formats:
   * 10/05/2026
   * 10/05/2026 à 21:13
   * 10-05-2026 21:13
   */
  const fr = text.match(
    /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s*(?:à)?\s*(\d{1,2}):(\d{2}))?$/
  );

  if (fr) {
    const [, d, m, y, hh = "0", mm = "0"] = fr;

    const dt = new Date(
      Number(y),
      Number(m) - 1,
      Number(d),
      Number(hh),
      Number(mm),
      0
    );

    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  // ISO / native parse fallback
  const native = new Date(text);

  return Number.isNaN(native.getTime()) ? null : native;
}

function stringifyCell(cell: RawCell): string {
  if (cell == null) return "";
  if (typeof cell === "boolean") return cell ? "oui" : "non";
  if (typeof cell === "number" && Number.isFinite(cell))
    return String(cell).replace(/\.0+$/, "").replace(/\.$/, "");
  return String(cell).trim();
}

function parseHeaderDateLabel(cell: RawCell): string | null {
  const raw = stringifyCell(cell);
  if (!raw) return null;

  const normalized = raw.replace(/\u00a0/g, " ");
  const match = normalized.match(
    /Date\s*['"]?\s*(?:de\s*)?(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+\d{1,2}:\d{2})?\s*(?:à|-)\s*(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+\d{1,2}:\d{2})?/i,
  );
  if (match) {
    const from = match[1];
    const to = match[2];
    return from === to ? from : `${from} — ${to}`;
  }

  const single = normalized.match(
    /Date\s*['"]?\s*(?:de\s*)?(\d{1,2}\.\d{1,2}\.\d{4})(?:\s+\d{1,2}:\d{2})?/i,
  );
  if (single) return single[1];

  const dates = normalized.match(/\d{1,2}\.\d{1,2}\.\d{4}/g);
  if (!dates?.length) return null;
  if (dates.length === 1) return dates[0];
  return dates[0] === dates[dates.length - 1]
    ? dates[0]
    : `${dates[0]} — ${dates[dates.length - 1]}`;
}

function omitCellA1(ws: XLSX.WorkSheet): XLSX.WorkSheet {
  const clone = { ...ws } as XLSX.WorkSheet;
  delete (clone as Record<string, unknown>).A1;
  return clone;
}

function rowOperationalHasSignal(r: Record<RowField, string>): boolean {
  return (
    !!(r.metier?.trim()) ||
    !!(r.regions?.trim()) ||
    !!(r.résultat?.trim()) ||
    !!(r.téléopérateur?.trim()) ||
    !!(r.natureRéclamation?.trim())
  );
}

function emptyFallback(v: string, isRegionPlaceholder: boolean): string {
  const t = v.trim();
  if (t) return t;
  if (isRegionPlaceholder) return FAUX_APPELS;
  return NON_RENSEIGNE;
}

export function parseWorkbook(wb: XLSX.WorkBook, fileHintName = ""): ParseResult {
  const logs: string[] = [];
  const sheets = wb.SheetNames;
  logs.push(`Fichier: ${fileHintName || "(buffer)"}`);
  logs.push(`Feuilles: ${sheets.join(", ")}`);
  const sheet = wb.Sheets[sheets[0]];
  const headerDateLabel = parseHeaderDateLabel((sheet?.A1 as { v?: RawCell })?.v ?? null) ?? undefined;
  const first = omitCellA1(sheet);
  const json = XLSX.utils.sheet_to_json<unknown[]>(first, {
    header: 1,
    blankrows: false,
    raw: false,
    defval: "",
  }) as RawCell[][];

  if (!json.length) {
    const debugEmpty: ParseDebug = {
      detectedHeaders: [],
      normalizedHeaders: [],
      totalRowsRaw: 0,
      parsedRows: 0,
      validRows: 0,
      invalidRows: 0,
      sheets,
      logs: [...logs, "Aucune donnée lisible dans la première feuille"],
    };
    return { rows: [], debug: debugEmpty };
  }

  const headerRowIdx = json.findIndex(
    (r) => Array.isArray(r) && r.some((c) => String(c ?? "").trim() !== ""),
  );
  const headerRow = json[headerRowIdx] ?? [];
  const headers = headerRow.map((h) =>
    nk(String(h ?? "")) === "" ? "" : String(h ?? "").trim(),
  );

  const colMap = buildHeaderMap(headers);
  const normalizedHeaders: HeaderMappingEntry[] = headers.map((h, i) => {
    const f = colMap.get(i);
    return {
      original: h,
      normalizedKey: f ?? "—",
      displayLabel: h ? displayLabel(h) : `Colonne ${i + 1}`,
    };
  });

  const rows: CrcRow[] = [];
  let validRows = 0;
  let invalidRows = 0;

  for (let r = headerRowIdx + 1; r < json.length; r++) {
    const line = json[r];
    if (!line) continue;
    const rec: Record<RowField, string> = {
      campagneNom: "",
      campagneType: "",
      formulaire: "",
      date: "",
      mois: "",
      téléopérateur: "",
      résultat: "",
      tempsAttente: "",
      tempsAttenteIvr: "",
      tempsAttenteQueue: "",
      téléopérateurNote: "",
      téléphone: "",
      adresse: "",
      nomPrénom: "",
      nIdentité: "",
      police: "",
      natureRéclamation: "",
      page3Type: "",
      regions: "",
      provinces: "",
      communes: "",
      metier: "",
    };

    colMap.forEach((field, colIdx) => {
      if (!field || field === "date" || field === "mois") return;
      const cell = line[colIdx];
      rec[field] = stringifyCell(cell);
    });

    const dateCol = [...colMap.entries()].find(([, f]) => f === "date")?.[0];
    const moisCol = [...colMap.entries()].find(([, f]) => f === "mois")?.[0];
    const rawDateCell =
      dateCol !== undefined ? (line[dateCol] as RawCell) : undefined;
    const rawMoisCell =
      moisCol !== undefined ? (line[moisCol] as RawCell) : undefined;

    const parsedDate = parseFlexibleDate(rawDateCell ?? "");
    let moisLabel = stringifyCell(rawMoisCell ?? "");
    if (!moisLabel && parsedDate) {
      moisLabel = `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(2, "0")}`;
    }

    const opOk = rowOperationalHasSignal(rec);
    const dateOk = parsedDate !== null && !Number.isNaN(+parsedDate);
    const valid = dateOk && opOk;

    const regionsRawTrim = rec.regions.trim();
    const régionCanon = regionsRawTrim
      ? regionToCanonical(regionsRawTrim)
      : FAUX_APPELS;

    if (valid) validRows++;
    else invalidRows++;

    const résultatTrim = rec.résultat.trim();
    const résultatRaw = résultatTrim;
    const résultatNormalized = !résultatTrim
      ? NON_RENSEIGNE
      : normalizeResult(résultatTrim) || NON_RENSEIGNE;

    const rowOut: CrcRow = {
      rawIndex: r,
      valid,
      validationReason:
        !dateOk ? "Date manquante ou invalide" : !opOk ? "Champs opérationnels absents" : undefined,
      campagneNom: emptyFallback(rec.campagneNom, false),
      campagneType: emptyFallback(rec.campagneType, false),
      formulaire: emptyFallback(rec.formulaire, false),
      date: parsedDate,
rawDateText: stringifyCell(rawDateCell ?? ""),
      moisLabel: moisLabel || NON_RENSEIGNE,
      téléopérateur: emptyFallback(rec.téléopérateur, false),
      résultat: résultatNormalized,
      résultatRaw,
      tempsAttente: emptyFallback(rec.tempsAttente, false),
      tempsAttenteIvr: emptyFallback(rec.tempsAttenteIvr, false),
      tempsAttenteQueue: emptyFallback(rec.tempsAttenteQueue, false),
      téléopérateurNote: emptyFallback(rec.téléopérateurNote, false),
      téléphone: emptyFallback(rec.téléphone, false),
      adresse: emptyFallback(rec.adresse, false),
      nomPrénom: emptyFallback(rec.nomPrénom, false),
      nIdentité: emptyFallback(rec.nIdentité, false),
      police: emptyFallback(rec.police, false),
      natureRéclamation: emptyFallback(rec.natureRéclamation, false),
      page3Type: emptyFallback(rec.page3Type, false),
      regions: regionsRawTrim ? regionsRawTrim : FAUX_APPELS,
      provinces: emptyFallback(rec.provinces, false),
      communes: emptyFallback(rec.communes, false),
      metier: emptyFallback(rec.metier, false),
      régionCanon,
    };

    rows.push(rowOut);
  }

  logs.push(`Lignes brutes données: ${json.length - headerRowIdx - 1}`);
  logs.push(`Lignes transformées (conservées intégralement): ${rows.length}`);

  const debug: ParseDebug = {
    detectedHeaders: headers.filter(Boolean),
    normalizedHeaders,
    totalRowsRaw: Math.max(json.length - headerRowIdx - 1, 0),
    parsedRows: rows.length,
    validRows,
    invalidRows,
    sheets,
    logs,
  };

  return { rows, debug, headerDateLabel };
}

/** Parse XLSX / XLS binary */
export function parseAxilusBuffer(buf: ArrayBuffer, fileHintName = ""): ParseResult {
  const wb = XLSX.read(buf, {
    type: "array",
    cellDates: true,
    dense: false,
  });
  return parseWorkbook(wb, fileHintName);
}

function sniffCsvDelimiter(sample: string): string {
  const firstLine = sample.split(/\r?\n/).find((l) => l.trim().length)?.trim() ?? "";
  const commas = (firstLine.match(/,/g) ?? []).length;
  const semis = (firstLine.match(/;/g) ?? []).length;
  return semis > commas ? ";" : ",";
}

export function decodeCsvBuffer(buf: ArrayBuffer): string {
  const utf = new TextDecoder("utf-8", { fatal: false }).decode(buf);
  if (!/\uFFFD/.test(utf)) return utf;
  return new TextDecoder("iso-8859-1").decode(buf);
}

export function parseAxilusCsvText(text: string, name = ""): ParseResult {
  const fs = sniffCsvDelimiter(text);
  const wb = XLSX.read(text, { type: "string", FS: fs, cellDates: true });
  return parseWorkbook(wb, name);
}

/** Public: route by filename */
export async function parseImportedFile(file: File): Promise<ParseResult> {
  const buf = await file.arrayBuffer();
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".csv")) return parseAxilusCsvText(decodeCsvBuffer(buf), file.name);
  return parseAxilusBuffer(buf, file.name);
}
