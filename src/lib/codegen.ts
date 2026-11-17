// Generates the code that reproduces the playground's current settings.
import { csvParseOptions, csvWriteOptions, readOptions } from "./engine";
import { DEFAULT_CSV_PARSE, DEFAULT_CSV_WRITE, DEFAULT_READ, type Settings } from "./state";

export const PACKAGE = "javascript-excel";

export type PackageManager = "npm" | "pnpm" | "yarn" | "bun";

/** The install command for the exact version the playground runs. */
export function installCommand(pm: PackageManager, version: string): string {
  const spec = `${PACKAGE}@${version}`;
  switch (pm) {
    case "npm":
      return `npm install ${spec}`;
    case "pnpm":
      return `pnpm add ${spec}`;
    case "yarn":
      return `yarn add ${spec}`;
    case "bun":
      return `bun add ${spec}`;
  }
}

/** A JS literal for an options object, leaving out keys equal to `defaults`. */
export function literal(value: Record<string, unknown>, defaults: Record<string, unknown> = {}, indent = 0): string {
  const entries = Object.entries(value).filter(([k, v]) => v !== undefined && v !== defaults[k]);
  if (entries.length === 0) return "{}";
  const pad = "  ".repeat(indent + 1);
  const end = "  ".repeat(indent);
  const lines = entries.map(([k, v]) => `${pad}${/^[A-Za-z_$][\w$]*$/.test(k) ? k : JSON.stringify(k)}: ${valueLiteral(v, indent + 1)},`);
  return `{\n${lines.join("\n")}\n${end}}`;
}

function valueLiteral(v: unknown, indent: number): string {
  if (typeof v === "string") return quote(v);
  if (v && typeof v === "object" && !Array.isArray(v)) return literal(v as Record<string, unknown>, {}, indent);
  return JSON.stringify(v);
}

/** A single-quoted JS string. */
export function quote(s: string): string {
  return `'${s.replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/\r/g, "\\r").replace(/\n/g, "\\n").replace(/\t/g, "\\t")}'`;
}

/** `(arg)` or `(arg, { ... })` when options differ from the defaults. */
function call(fn: string, arg: string, opts: string): string {
  return opts === "{}" ? `${fn}(${arg})` : `${fn}(${arg}, ${opts})`;
}

const WORKER_FILE = `// excel.worker.ts
import { exposeWorker } from 'javascript-excel/worker';
exposeWorker();
`;

const WORKER_CLIENT = `import { createWorkerClient } from 'javascript-excel/worker';

const excel = createWorkerClient(
  new Worker(new URL('./excel.worker.ts', import.meta.url), { type: 'module' }),
);`;

export interface Snippet {
  /** Files to create, in order; the main file comes last. */
  files: Array<{ name: string; code: string }>;
}

/** Code for the Read tab: open a user's .xlsx (or .csv) and walk the result. */
export function readSnippet(s: Settings, fileKind: "xlsx" | "csv" = "xlsx"): Snippet {
  if (fileKind === "csv") {
    const opts = literal(csvParseOptions(s.csvParse) as Record<string, unknown>, { ...DEFAULT_CSV_PARSE } as Record<string, unknown>);
    const body = s.worker
      ? `${WORKER_CLIENT}

export async function openCsv(file: File) {
  const { rows, delimiter, warnings } = await ${call("excel.parseCsv", "file", opts)};
  console.log(\`\${rows.length} rows, delimiter \${JSON.stringify(delimiter)}\`);
  for (const w of warnings) console.warn(w.code, w.ref, w.message);
  return rows;
}
`
      : `import { readCsv } from 'javascript-excel/csv';

export async function openCsv(file: File) {
  const { rows, delimiter, warnings } = await ${call("readCsv", "file", opts)};
  console.log(\`\${rows.length} rows, delimiter \${JSON.stringify(delimiter)}\`);
  for (const w of warnings) console.warn(w.code, w.ref, w.message);
  return rows;
}
`;
    return { files: [...(s.worker ? [{ name: "excel.worker.ts", code: WORKER_FILE }] : []), { name: "open-csv.ts", code: body }] };
  }
  const opts = literal(readOptions(s.read) as Record<string, unknown>, { ...DEFAULT_READ } as Record<string, unknown>);
  const imports = s.worker
    ? `${WORKER_CLIENT}
import { formatCell } from 'javascript-excel/format';
import { createRenderSheet, renderCell, sheetExtent } from 'javascript-excel/render';`
    : `import { read } from 'javascript-excel/read';
import { formatCell } from 'javascript-excel/format';
import { createRenderSheet, renderCell, sheetExtent } from 'javascript-excel/render';`;
  const readCall = call(s.worker ? "excel.read" : "read", "file", opts);
  const locale = s.read.locale === "en-US" ? "" : `, locale: ${quote(s.read.locale)}`;
  const body = `${imports}

export async function openWorkbook(file: File) {
  const { workbook, warnings } = await ${readCall};
  for (const w of warnings) console.warn(w.code, w.sheet, w.ref, w.message);

  for (const sheet of workbook.sheets) {
    console.log(sheet.name, sheet.dimension, 'merges:', sheet.merges ?? []);
  }

  // Displayed text, concrete styles and merges, as a viewer draws them.
  const view = createRenderSheet(workbook, workbook.activeSheet ?? 0, { format: formatCell${locale} });
  const { rows, cols } = sheetExtent(view);
  for (let r = 0; r < Math.min(rows, 20); r++) {
    for (let c = 0; c < Math.min(cols, 10); c++) {
      const cell = renderCell(view, r, c);
      if (cell?.text) console.log(r, c, cell.text, cell.style.font.bold ? '(bold)' : '');
    }
  }
  return workbook;
}
`;
  return { files: [...(s.worker ? [{ name: "excel.worker.ts", code: WORKER_FILE }] : []), { name: "open-workbook.ts", code: body }] };
}

/** Code for the Write tab: a dataset to .xlsx and .csv downloads. */
export function writeSnippet(s: Settings): Snippet {
  const w = s.write;
  const opts: string[] = [`  name: ${quote(w.sheetName)},`];
  if (w.freezeHeader) opts.push("  freezeHeader: true,");
  if (w.autoFilter) opts.push("  autoFilter: true,");
  if (w.autoWidth) opts.push("  autoWidth: true,");
  opts.push(`  headerStyle: { font: { bold: true }${w.headerFill ? `, fill: ${quote(w.headerFill)}` : ""} },`);
  if (w.totals) opts.push("  totals: { label: 'Total', style: { font: { bold: true } } },");
  if (w.banded) opts.push("  rowStyle: (_row, i) => (i % 2 ? stripe : undefined),");
  const writeOpts = w.compression === "deflate" ? "" : ", { compression: 'store' }";
  const imports = s.worker
    ? `import { fromDataset, type Dataset } from 'javascript-excel';
import { toCsvBlob } from 'javascript-excel/csv';
import { download } from 'javascript-excel/dom';
${WORKER_CLIENT}`
    : `import { fromDataset, toBlob, type Dataset } from 'javascript-excel';
import { toCsvBlob } from 'javascript-excel/csv';
import { download } from 'javascript-excel/dom';`;
  const xlsx = s.worker
    ? `  const bytes = await excel.write({ sheets: [sheet] }${writeOpts});
  download(bytes, 'dataset.xlsx');`
    : `  download(await toBlob({ sheets: [sheet] }${writeOpts}), 'dataset.xlsx');`;
  const csvOpts = literal(csvWriteOptions(s.csvWrite) as Record<string, unknown>, { ...DEFAULT_CSV_WRITE } as Record<string, unknown>);
  const body = `${imports}

${w.banded ? "const stripe = { fill: 'F2F2F2' };\n\n" : ""}export async function exportDataset(dataset: Dataset<unknown>) {
  const sheet = fromDataset(dataset, {
${opts.join("\n")}
  });
${xlsx}
  download(${call("toCsvBlob", "sheet", csvOpts)}, 'dataset.csv');
}
`;
  return { files: [...(s.worker ? [{ name: "excel.worker.ts", code: WORKER_FILE }] : []), { name: "export-dataset.ts", code: body }] };
}

/** Code for the CSV tab: parse text, then write it back with the chosen options. */
export function csvSnippet(s: Settings): Snippet {
  const { encoding: _enc, ...parse } = csvParseOptions(s.csvParse);
  void _enc;
  const parseOpts = literal(parse as Record<string, unknown>, { ...DEFAULT_CSV_PARSE } as Record<string, unknown>);
  const writeOpts = literal(csvWriteOptions(s.csvWrite) as Record<string, unknown>, { ...DEFAULT_CSV_WRITE, bom: false } as Record<string, unknown>);
  const body = s.worker
    ? `${WORKER_CLIENT}

export async function roundTrip(text: string) {
  const { rows, delimiter, newline, warnings } = await ${call("excel.parseCsv", "text", parseOpts)};
  for (const w of warnings) console.warn(w.code, w.ref, w.message);
  console.log({ delimiter, newline, rows });
  return excel.toCsv(rows${writeOpts === "{}" ? "" : `, ${writeOpts}`});
}
`
    : `import { parseCsv, toCsv } from 'javascript-excel/csv';

export function roundTrip(text: string) {
  const { rows, delimiter, newline, warnings } = ${call("parseCsv", "text", parseOpts)};
  for (const w of warnings) console.warn(w.code, w.ref, w.message);
  console.log({ delimiter, newline, rows });
  return ${call("toCsv", "rows", writeOpts)};
}
`;
  return { files: [...(s.worker ? [{ name: "excel.worker.ts", code: WORKER_FILE }] : []), { name: "csv-round-trip.ts", code: body }] };
}

export function snippetFor(s: Settings, fileKind: "xlsx" | "csv" = "xlsx"): Snippet {
  return s.tab === "write" ? writeSnippet(s) : s.tab === "csv" ? csvSnippet(s) : readSnippet(s, fileKind);
}

/** All files of a snippet as one copyable text. */
export function snippetText(snippet: Snippet): string {
  return snippet.files.map((f) => (snippet.files.length > 1 ? `// ---- ${f.name}\n${f.code}` : f.code)).join("\n");
}
