// Calls into javascript-excel, on the main thread or in a Web Worker, timed.
import { fromDataset, fromDatasets, isExcelError, toUint8Array, type CellInput, type Dataset, type DatasetBlock, type DatasetOptions, type ExcelWarning, type Progress, type Sheet, type Style, type Workbook } from "javascript-excel";
import { parseCsv, readCsv, toCsv, type CsvParseOptions, type CsvParseResult, type CsvWriteOptions } from "javascript-excel/csv";
import { read, type ReadOptions } from "javascript-excel/read";
import { createWorkerClient, type WorkerClient } from "javascript-excel/worker";
import type { CsvParseSettings, CsvWriteSettings, ReadSettings, WriteSettings } from "./state";

let client: WorkerClient | undefined;
function worker(): WorkerClient {
  client ??= createWorkerClient(new Worker(new URL("../excel.worker.ts", import.meta.url), { type: "module" }));
  return client;
}

export interface Timed<T> {
  result: T;
  ms: number;
  where: "main" | "worker";
}

async function timed<T>(useWorker: boolean, run: () => Promise<T>): Promise<Timed<T>> {
  const started = performance.now();
  const result = await run();
  return { result, ms: performance.now() - started, where: useWorker ? "worker" : "main" };
}

export const isCsvName = (name: string): boolean => /\.(csv|tsv|txt)$/i.test(name);

export function readOptions(s: ReadSettings): ReadOptions {
  const o: ReadOptions = { styles: s.styles, formulas: s.formulas, dates: s.dates, richText: s.richText };
  if (s.maxRows > 0) o.maxRows = s.maxRows;
  return o;
}

export function csvParseOptions(s: CsvParseSettings): CsvParseOptions & { encoding?: string } {
  const o: CsvParseOptions & { encoding?: string } = { delimiter: s.delimiter, inferTypes: s.inferTypes, skipEmptyLines: s.skipEmptyLines, mode: s.mode };
  if (s.encoding !== "utf-8") o.encoding = s.encoding;
  return o;
}

export function csvWriteOptions(s: CsvWriteSettings): CsvWriteOptions {
  return { delimiter: s.delimiter, newline: s.newline, quote: s.quote, bom: s.bom, escapeFormulas: s.escapeFormulas };
}

export interface Loaded {
  kind: "xlsx" | "csv";
  name: string;
  size: number;
  workbook: Workbook<CellInput>;
  warnings: ExcelWarning[];
  csv?: Pick<CsvParseResult, "delimiter" | "newline" | "bom">;
}

/** Reads an .xlsx or .csv file into a workbook. */
export async function readFile(
  name: string,
  data: Uint8Array | Blob,
  settings: { read: ReadSettings; csvParse: CsvParseSettings; worker: boolean },
  extra: { signal?: AbortSignal; onProgress?: (p: Progress) => void } = {},
): Promise<Timed<Loaded>> {
  const size = data instanceof Blob ? data.size : data.byteLength;
  const useWorker = settings.worker;
  if (isCsvName(name)) {
    const opts = { ...csvParseOptions(settings.csvParse), ...(extra.signal ? { signal: extra.signal } : {}) };
    return timed(useWorker, async () => {
      const r = useWorker ? await worker().parseCsv(data, opts) : await readCsv(data, opts);
      const sheet: Sheet<CellInput> = { name: sheetNameFor(name), rows: r.rows };
      return { kind: "csv", name, size, workbook: { sheets: [sheet] }, warnings: r.warnings, csv: { delimiter: r.delimiter, newline: r.newline, bom: r.bom } };
    });
  }
  const opts: ReadOptions = { ...readOptions(settings.read), ...extra };
  return timed(useWorker, async () => {
    const r = useWorker ? await worker().read(data, opts) : await read(data, opts);
    return { kind: "xlsx", name, size, workbook: r.workbook, warnings: r.warnings };
  });
}

function sheetNameFor(file: string): string {
  return (
    file
      .replace(/\.[^.]+$/, "")
      .replace(/[\\/?*[\]:]/g, " ")
      .replace(/^'+|'+$/g, "")
      .trim()
      .slice(0, 31) || "CSV"
  );
}

/** Parses CSV text synchronously (main thread) or in the worker. */
export function parseCsvText(text: string, s: CsvParseSettings, useWorker: boolean): Promise<Timed<CsvParseResult>> {
  const { encoding: _encoding, ...opts } = csvParseOptions(s);
  void _encoding; // text is already decoded
  return timed(useWorker, async () => (useWorker ? worker().parseCsv(text, opts) : parseCsv(text, opts)));
}

export function csvTextOf(rows: CellInput[][], s: CsvWriteSettings, useWorker: boolean): Promise<Timed<string>> {
  const opts = csvWriteOptions(s);
  return timed(useWorker, async () => (useWorker ? worker().toCsv(rows, opts) : toCsv(rows, opts)));
}

export function datasetOptions(s: WriteSettings, onWarning?: (w: ExcelWarning) => void): DatasetOptions<unknown> {
  const headerStyle: Style = { font: { bold: true }, ...(s.headerFill ? { fill: s.headerFill } : {}) };
  const o: DatasetOptions<unknown> = { name: s.sheetName, freezeHeader: s.freezeHeader, autoFilter: s.autoFilter, autoWidth: s.autoWidth, headerStyle };
  if (s.totals) o.totals = { label: "Total", style: { font: { bold: true } } };
  if (s.banded) {
    const stripe: Style = { fill: "F2F2F2" };
    o.rowStyle = (_row, i) => (i % 2 ? stripe : undefined);
  }
  if (onWarning) o.onWarning = onWarning;
  return o;
}

export interface Built {
  sheet: Sheet;
  warnings: ExcelWarning[];
  /** 1 for a single dataset (fromDataset), n for an array of blocks (fromDatasets). */
  blocks: number;
}

/** The blocks of an array of datasets, with `placement` set on every block after the first. */
export function placeBlocks(datasets: Array<Dataset<unknown>>, placement: WriteSettings["placement"]): Array<DatasetBlock<unknown>> {
  return datasets.map((d, i) => (i === 0 || placement === "below" ? d : { ...d, placement }));
}

/** Parses the dataset JSON and builds a sheet. Throws a readable Error on bad JSON. */
export function buildSheet(json: string, s: WriteSettings): Timed<Built> {
  let dataset: Dataset<unknown> | Array<Dataset<unknown>>;
  try {
    dataset = JSON.parse(json) as Dataset<unknown> | Array<Dataset<unknown>>;
  } catch (e) {
    throw new Error(`The dataset is not valid JSON: ${(e as Error).message}`, { cause: e });
  }
  const started = performance.now();
  const warnings: ExcelWarning[] = [];
  const options = datasetOptions(s, (w) => warnings.push(w));
  if (Array.isArray(dataset)) {
    const sheet = fromDatasets(placeBlocks(dataset, s.placement), { ...options, gap: s.gap });
    return { result: { sheet, warnings, blocks: dataset.length }, ms: performance.now() - started, where: "main" };
  }
  const sheet = fromDataset(dataset, options);
  return { result: { sheet, warnings, blocks: 1 }, ms: performance.now() - started, where: "main" };
}

export function writeXlsx(workbook: Workbook, s: WriteSettings, useWorker: boolean): Promise<Timed<Uint8Array>> {
  const opts = { compression: s.compression };
  return timed(useWorker, async () => (useWorker ? worker().write(workbook, opts) : toUint8Array(workbook, opts)));
}

/** Human-readable error text, with the library's error code when there is one. */
export function describeError(e: unknown): string {
  if (isExcelError(e)) return `${e.code}: ${e.message}`;
  if (e instanceof Error) return e.name === "AbortError" ? "Cancelled." : e.message;
  return String(e);
}

export function formatMs(ms: number): string {
  return ms < 10 ? `${ms.toFixed(2)} ms` : ms < 1000 ? `${ms.toFixed(1)} ms` : `${(ms / 1000).toFixed(2)} s`;
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KiB`;
  return `${(n / 1024 / 1024).toFixed(1)} MiB`;
}
