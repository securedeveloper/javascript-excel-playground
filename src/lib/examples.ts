// One-click scenarios: each one sets the tab and settings, and says what to look at.
import { CSV_SAMPLES, SAMPLE_BLOCKS_JSON, SAMPLE_DATASET, SAMPLE_DATASET_JSON } from "./samples";
import type { Settings } from "./state";

export interface Example {
  id: string;
  title: string;
  description: string;
  apply: (s: Settings) => Settings;
}

const csv = (id: string): string => CSV_SAMPLES.find((c) => c.id === id)?.text ?? "";

export const EXAMPLES: readonly Example[] = [
  {
    id: "read-sample",
    title: "Read the sample workbook",
    description: "Styles, number formats, merges, formulas, links and a hidden sheet.",
    apply: (s) => ({ ...s, tab: "read", read: { ...s.read, styles: true, formulas: true, dates: "date" } }),
  },
  {
    id: "read-serial",
    title: "Exact dates (serials) in de-DE",
    description: "dates: 'serial' with formatCell renders dates as Excel does, in German.",
    apply: (s) => ({ ...s, tab: "read", read: { ...s.read, dates: "serial", locale: "de-DE" } }),
  },
  {
    id: "read-fast",
    title: "Fast preview read",
    description: "styles: false and maxRows: 5 — the shape for quick previews (ROWS_TRUNCATED).",
    apply: (s) => ({ ...s, tab: "read", read: { ...s.read, styles: false, maxRows: 5 } }),
  },
  {
    id: "read-worker",
    title: "Read in a Web Worker",
    description: "The same read, off the main thread; the workbook comes back by structured clone.",
    apply: (s) => ({ ...s, tab: "read", worker: true }),
  },
  {
    id: "write-groups",
    title: "Dataset with header groups and totals",
    description: "Nested columns become merged header rows; total: 'sum' adds SUM formulas.",
    apply: (s) => ({ ...s, tab: "write", dataset: SAMPLE_DATASET_JSON, write: { ...s.write, totals: true, freezeHeader: true, autoFilter: true } }),
  },
  {
    id: "write-side-by-side",
    title: "Two blocks side by side",
    description: "An array of datasets goes through fromDatasets; placement: 'right' puts the summary next to the table.",
    apply: (s) => ({ ...s, tab: "write", dataset: SAMPLE_BLOCKS_JSON, write: { ...s.write, placement: "right", gap: 1, autoFilter: false } }),
  },
  {
    id: "write-banded",
    title: "Banded rows, no compression",
    description: "rowStyle stripes every other row; compression: 'store' writes faster, bigger files.",
    apply: (s) => ({
      ...s,
      tab: "write",
      dataset: JSON.stringify({ ...SAMPLE_DATASET, data: [...SAMPLE_DATASET.data, ...SAMPLE_DATASET.data.map((d) => ({ ...d, name: `${d.name} II` }))] }, null, 2),
      write: { ...s.write, banded: true, headerFill: "E2EFDA", compression: "store" },
    }),
  },
  {
    id: "csv-semicolon",
    title: "European CSV",
    description: "Delimiter detection picks ';', quoted decimals stay text, multi-line fields survive.",
    apply: (s) => ({ ...s, tab: "csv", csvText: csv("semicolon"), csvParse: { ...s.csvParse, delimiter: "auto", inferTypes: true, mode: "lenient" }, csvWrite: { ...s.csvWrite, delimiter: ";" } }),
  },
  {
    id: "csv-messy",
    title: "Messy CSV, lenient vs strict",
    description: "Lenient mode keeps going with warnings; switch to strict to see CSV_MALFORMED.",
    apply: (s) => ({ ...s, tab: "csv", csvText: csv("messy"), csvParse: { ...s.csvParse, mode: "lenient", skipEmptyLines: false } }),
  },
  {
    id: "csv-injection",
    title: "Formula-injection guard",
    description: "Text starting with = + - @ is written with a leading ' unless escapeFormulas is off.",
    apply: (s) => ({ ...s, tab: "csv", csvText: csv("messy"), csvParse: { ...s.csvParse, mode: "lenient" }, csvWrite: { ...s.csvWrite, escapeFormulas: true, quote: "all" } }),
  },
];
