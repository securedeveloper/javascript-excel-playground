// The playground's settings: what the URL, the history and the code generator
// all describe. File contents are never part of it (they stay in memory only).

export type Tab = "read" | "write" | "csv";
export const TABS: readonly Tab[] = ["read", "write", "csv"];

export type Delimiter = "auto" | "," | ";" | "\t" | "|";

export interface ReadSettings {
  styles: boolean;
  formulas: boolean;
  dates: "date" | "serial";
  richText: "runs" | "plain";
  /** 0 = no limit. */
  maxRows: number;
  /** Locale for displayed (number-formatted) text in the grid. */
  locale: string;
}

export interface CsvParseSettings {
  delimiter: Delimiter;
  inferTypes: boolean;
  skipEmptyLines: boolean;
  mode: "lenient" | "strict";
  encoding: string;
}

export interface CsvWriteSettings {
  delimiter: "," | ";" | "\t" | "|";
  newline: "\r\n" | "\n";
  quote: "auto" | "all";
  bom: boolean;
  escapeFormulas: boolean;
}

export interface WriteSettings {
  sheetName: string;
  freezeHeader: boolean;
  autoFilter: boolean;
  autoWidth: boolean;
  totals: boolean;
  /** RRGGBB, or "" for no header fill. */
  headerFill: string;
  banded: boolean;
  compression: "deflate" | "store";
}

export interface Settings {
  tab: Tab;
  worker: boolean;
  read: ReadSettings;
  csvParse: CsvParseSettings;
  csvWrite: CsvWriteSettings;
  write: WriteSettings;
  /** The dataset JSON of the Write tab. */
  dataset: string;
  /** The CSV text of the CSV tab. */
  csvText: string;
}

export const LOCALES = ["en-US", "en-GB", "de-DE", "fr-FR", "ja-JP"] as const;

export const DEFAULT_READ: ReadSettings = { styles: true, formulas: true, dates: "date", richText: "runs", maxRows: 0, locale: "en-US" };
export const DEFAULT_CSV_PARSE: CsvParseSettings = { delimiter: "auto", inferTypes: true, skipEmptyLines: false, mode: "lenient", encoding: "utf-8" };
export const DEFAULT_CSV_WRITE: CsvWriteSettings = { delimiter: ",", newline: "\r\n", quote: "auto", bom: true, escapeFormulas: true };
export const DEFAULT_WRITE: WriteSettings = {
  sheetName: "Dataset",
  freezeHeader: true,
  autoFilter: true,
  autoWidth: true,
  totals: true,
  headerFill: "DDEBF7",
  banded: false,
  compression: "deflate",
};

export function defaultSettings(dataset: string, csvText: string): Settings {
  return {
    tab: "read",
    worker: false,
    read: { ...DEFAULT_READ },
    csvParse: { ...DEFAULT_CSV_PARSE },
    csvWrite: { ...DEFAULT_CSV_WRITE },
    write: { ...DEFAULT_WRITE },
    dataset,
    csvText,
  };
}

/** Longest dataset or CSV text kept in a shared URL, in characters. */
export const MAX_URL_TEXT = 4000;

type Plain = Record<string, unknown>;

/** The fields of `value` that differ from `base`, one level deep for nested objects. */
function diff(value: Plain, base: Plain): Plain {
  const out: Plain = {};
  for (const [k, v] of Object.entries(value)) {
    const b = base[k];
    if (v && typeof v === "object" && b && typeof b === "object") {
      const d = diff(v as Plain, b as Plain);
      if (Object.keys(d).length) out[k] = d;
    } else if (v !== b) out[k] = v;
  }
  return out;
}

function toBase64Url(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function fromBase64Url(s: string): string {
  const b64 = s.replace(/-/g, "+").replace(/_/g, "/");
  const bin = atob(b64 + "=".repeat((4 - (b64.length % 4)) % 4));
  return new TextDecoder().decode(Uint8Array.from(bin, (c) => c.charCodeAt(0)));
}

/** The query string for `settings`: only what differs from `defaults`. Long texts are left out. */
export function encodeSettings(settings: Settings, defaults: Settings): string {
  const d = diff(settings as unknown as Plain, defaults as unknown as Plain);
  for (const key of ["dataset", "csvText"] as const) {
    if (typeof d[key] === "string" && (d[key] as string).length > MAX_URL_TEXT) delete d[key];
  }
  const params = new URLSearchParams();
  if (d.tab) params.set("tab", String(d.tab));
  delete d.tab;
  if (Object.keys(d).length) params.set("s", toBase64Url(JSON.stringify(d)));
  return params.toString();
}

const pick = <T>(value: unknown, allowed: readonly T[], fallback: T): T => (allowed.includes(value as T) ? (value as T) : fallback);
const bool = (value: unknown, fallback: boolean): boolean => (typeof value === "boolean" ? value : fallback);
const str = (value: unknown, fallback: string, max = 100_000): string => (typeof value === "string" && value.length <= max ? value : fallback);

/**
 * Settings from a query string, validated field by field against `defaults`:
 * a URL is user input, so anything unexpected falls back to the default.
 */
export function decodeSettings(query: string, defaults: Settings): Settings {
  const params = new URLSearchParams(query);
  let raw: Plain = {};
  const s = params.get("s");
  if (s) {
    try {
      const parsed: unknown = JSON.parse(fromBase64Url(s));
      if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) raw = parsed as Plain;
    } catch {
      raw = {};
    }
  }
  const obj = (v: unknown): Plain => (v && typeof v === "object" && !Array.isArray(v) ? (v as Plain) : {});
  const r = obj(raw.read);
  const cp = obj(raw.csvParse);
  const cw = obj(raw.csvWrite);
  const w = obj(raw.write);
  const D = defaults;
  const maxRows = typeof r.maxRows === "number" && Number.isInteger(r.maxRows) && r.maxRows >= 0 && r.maxRows <= 1_048_576 ? r.maxRows : D.read.maxRows;
  const fill = typeof w.headerFill === "string" && /^([0-9a-fA-F]{6})?$/.test(w.headerFill) ? w.headerFill.toUpperCase() : D.write.headerFill;
  return {
    tab: pick(params.get("tab"), TABS, D.tab),
    worker: bool(raw.worker, D.worker),
    read: {
      styles: bool(r.styles, D.read.styles),
      formulas: bool(r.formulas, D.read.formulas),
      dates: pick(r.dates, ["date", "serial"] as const, D.read.dates),
      richText: pick(r.richText, ["runs", "plain"] as const, D.read.richText),
      maxRows,
      locale: pick(r.locale, LOCALES as readonly string[], D.read.locale),
    },
    csvParse: {
      delimiter: pick(cp.delimiter, ["auto", ",", ";", "\t", "|"] as const, D.csvParse.delimiter),
      inferTypes: bool(cp.inferTypes, D.csvParse.inferTypes),
      skipEmptyLines: bool(cp.skipEmptyLines, D.csvParse.skipEmptyLines),
      mode: pick(cp.mode, ["lenient", "strict"] as const, D.csvParse.mode),
      encoding: pick(cp.encoding, ["utf-8", "utf-16le", "windows-1252", "iso-8859-1"] as const, D.csvParse.encoding as "utf-8"),
    },
    csvWrite: {
      delimiter: pick(cw.delimiter, [",", ";", "\t", "|"] as const, D.csvWrite.delimiter),
      newline: pick(cw.newline, ["\r\n", "\n"] as const, D.csvWrite.newline),
      quote: pick(cw.quote, ["auto", "all"] as const, D.csvWrite.quote),
      bom: bool(cw.bom, D.csvWrite.bom),
      escapeFormulas: bool(cw.escapeFormulas, D.csvWrite.escapeFormulas),
    },
    write: {
      sheetName: str(w.sheetName, D.write.sheetName, 31),
      freezeHeader: bool(w.freezeHeader, D.write.freezeHeader),
      autoFilter: bool(w.autoFilter, D.write.autoFilter),
      autoWidth: bool(w.autoWidth, D.write.autoWidth),
      totals: bool(w.totals, D.write.totals),
      headerFill: fill,
      banded: bool(w.banded, D.write.banded),
      compression: pick(w.compression, ["deflate", "store"] as const, D.write.compression),
    },
    dataset: str(raw.dataset, D.dataset),
    csvText: str(raw.csvText, D.csvText),
  };
}
