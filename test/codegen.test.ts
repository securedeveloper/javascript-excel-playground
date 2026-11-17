import * as core from "javascript-excel";
import * as csv from "javascript-excel/csv";
import * as format from "javascript-excel/format";
import * as readMod from "javascript-excel/read";
import * as render from "javascript-excel/render";
import ts from "typescript";
import { describe, expect, it } from "vitest";
import { csvSnippet, installCommand, readSnippet, snippetFor, writeSnippet, type Snippet } from "../src/lib/codegen";
import { buildSheet } from "../src/lib/engine";
import { EXAMPLES } from "../src/lib/examples";
import { CSV_SAMPLES, SAMPLE_DATASET, SAMPLE_DATASET_JSON, sampleXlsx } from "../src/lib/samples";
import { defaultSettings, type Settings } from "../src/lib/state";

const BASE = defaultSettings(SAMPLE_DATASET_JSON, CSV_SAMPLES[0]!.text);
const ENTRY_POINTS = new Set(["javascript-excel", "javascript-excel/read", "javascript-excel/csv", "javascript-excel/dom", "javascript-excel/worker", "javascript-excel/format", "javascript-excel/render", "javascript-excel/import"]);

/** Every file must be syntactically valid TypeScript and import only public entry points. */
function expectValid(snippet: Snippet): void {
  expect(snippet.files.length).toBeGreaterThan(0);
  for (const f of snippet.files) {
    const out = ts.transpileModule(f.code, { reportDiagnostics: true, fileName: f.name, compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 } });
    expect(out.diagnostics?.map((d) => ts.flattenDiagnosticMessageText(d.messageText, "\n")) ?? [], f.name).toEqual([]);
    for (const m of f.code.matchAll(/from '([^']+)'/g)) expect(ENTRY_POINTS.has(m[1]!), m[1]).toBe(true);
  }
}

/** Runs a snippet's main file as CommonJS against the real library. */
function load(snippet: Snippet): Record<string, (...args: unknown[]) => unknown> {
  const main = snippet.files.at(-1)!;
  const js = ts.transpileModule(main.code, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText;
  const modules: Record<string, unknown> = { "javascript-excel": core, "javascript-excel/csv": csv, "javascript-excel/read": readMod, "javascript-excel/format": format, "javascript-excel/render": render };
  const exports: Record<string, (...args: unknown[]) => unknown> = {};
  const quiet = { log: () => {}, warn: () => {} };
  new Function("require", "exports", "console", js)((id: string) => modules[id], exports, quiet);
  return exports;
}

const variants: Array<[string, Settings]> = [
  ["defaults", BASE],
  ["worker", { ...BASE, worker: true }],
  ["read options", { ...BASE, read: { ...BASE.read, styles: false, dates: "serial", maxRows: 10, locale: "de-DE" } }],
  ["csv options", { ...BASE, tab: "csv", csvParse: { ...BASE.csvParse, delimiter: "\t", mode: "strict" }, csvWrite: { ...BASE.csvWrite, delimiter: ";", quote: "all", newline: "\n" } }],
  ["write options", { ...BASE, tab: "write", write: { ...BASE.write, banded: true, headerFill: "", compression: "store", sheetName: "It's mine" } }],
  ...EXAMPLES.map((e) => [`example ${e.id}`, e.apply(BASE)] as [string, Settings]),
];

describe("code generator", () => {
  it.each(variants)("outputs valid TypeScript for every tab (%s)", (_name, s) => {
    for (const tab of ["read", "write", "csv"] as const) expectValid(snippetFor({ ...s, tab }));
    expectValid(readSnippet(s, "csv"));
  });

  it("pins the exact version in every install command", () => {
    expect(installCommand("npm", "2.0.0-alpha.7")).toBe("npm install javascript-excel@2.0.0-alpha.7");
    expect(installCommand("pnpm", "1.2.3")).toBe("pnpm add javascript-excel@1.2.3");
    expect(installCommand("yarn", "1.2.3")).toBe("yarn add javascript-excel@1.2.3");
    expect(installCommand("bun", "1.2.3")).toBe("bun add javascript-excel@1.2.3");
  });

  it("only spells out options that differ from the library defaults", () => {
    const plain = readSnippet(BASE).files.at(-1)!.code;
    expect(plain).toContain("await read(file)");
    const tuned = readSnippet({ ...BASE, read: { ...BASE.read, dates: "serial", maxRows: 25 } }).files.at(-1)!.code;
    expect(tuned).toContain("dates: 'serial'");
    expect(tuned).toContain("maxRows: 25");
    expect(tuned).not.toContain("styles:");
  });

  it("adds the worker file and client in worker mode", () => {
    const s = readSnippet({ ...BASE, worker: true });
    expect(s.files.map((f) => f.name)).toEqual(["excel.worker.ts", "open-workbook.ts"]);
    expect(s.files[0]!.code).toContain("exposeWorker()");
    expect(s.files[1]!.code).toContain("createWorkerClient");
    expect(s.files[1]!.code).toContain("excel.read(file)");
  });

  it("generates a CSV round trip that runs and honours the options", () => {
    const s: Settings = { ...BASE, tab: "csv", csvWrite: { ...BASE.csvWrite, delimiter: ";", newline: "\n", bom: false } };
    const { roundTrip } = load(csvSnippet(s));
    expect(roundTrip!("a,b\n1,=2\n")).toBe("a;b\n1;'=2\n");
  });

  it("generates a read snippet that runs on the sample workbook", async () => {
    const { openWorkbook } = load(readSnippet(BASE));
    const bytes = await sampleXlsx();
    const wb = (await openWorkbook!(new Blob([bytes as BlobPart]))) as core.Workbook;
    expect(wb.sheets.map((x) => x.name)).toEqual(["Sales", "Notes", "Hidden"]);
  });

  it("generates fromDatasets with placement for an array of datasets", async () => {
    const { SAMPLE_BLOCKS, SAMPLE_BLOCKS_JSON } = await import("../src/lib/samples");
    const s: Settings = { ...BASE, tab: "write", dataset: SAMPLE_BLOCKS_JSON, write: { ...BASE.write, placement: "right", gap: 2 } };
    const snippet = writeSnippet(s);
    expectValid(snippet);
    const code = snippet.files.at(-1)!.code;
    expect(code).toContain("fromDatasets(blocks");
    expect(code).toContain("placement: 'right'");
    expect(code).toContain("gap: 2,");
    expect(writeSnippet({ ...s, write: { ...s.write, placement: "below", gap: 1 } }).files.at(-1)!.code).not.toMatch(/placement:|gap:/);
    // The generated function builds the same sheet as the playground.
    const calls: unknown[] = [];
    const js = ts.transpileModule(code, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText;
    const exports: Record<string, (...a: unknown[]) => Promise<unknown>> = {};
    const written: core.Workbook[] = [];
    const fake = { ...core, toBlob: async (wb: core.Workbook) => (written.push(wb), new Blob([])) };
    new Function("require", "exports", js)((id: string) => ({ "javascript-excel": fake, "javascript-excel/csv": csv, "javascript-excel/dom": { download: (d: unknown) => calls.push(d) } })[id], exports);
    await exports.exportDatasets!(SAMPLE_BLOCKS);
    expect(calls).toHaveLength(2);
    expect(written[0]!.sheets[0]).toEqual(buildSheet(SAMPLE_BLOCKS_JSON, s.write).result.sheet);
  });

  it("generates a write snippet whose options build the same sheet", () => {
    const code = writeSnippet({ ...BASE, write: { ...BASE.write, sheetName: "It's mine" } }).files.at(-1)!.code;
    expect(code).toContain("name: 'It\\'s mine'");
    expect(code).toContain("freezeHeader: true");
    expect(code).toContain("total");
    expect(JSON.parse(SAMPLE_DATASET_JSON)).toEqual(SAMPLE_DATASET);
  });
});
