import { cellText, type Cell } from "javascript-excel";
import { formatCell } from "javascript-excel/format";
import { createRenderSheet, mergeAt, renderCell, sheetSettings } from "javascript-excel/render";
import { describe, expect, it } from "vitest";
import { buildSheet, parseCsvText, readFile } from "../src/lib/engine";
import { CSV_SAMPLES, SAMPLE_DATASET_JSON, SAMPLE_XLSX_NAME, sampleXlsx } from "../src/lib/samples";
import { defaultSettings } from "../src/lib/state";

const S = defaultSettings(SAMPLE_DATASET_JSON, CSV_SAMPLES[0]!.text);

describe("read flow on the built-in sample", () => {
  it("reads sheets, cells, styles, merges and links without warnings", async () => {
    const t = await readFile(SAMPLE_XLSX_NAME, await sampleXlsx(), S);
    expect(t.where).toBe("main");
    expect(t.ms).toBeGreaterThanOrEqual(0);
    const { workbook, warnings, kind } = t.result;
    expect(kind).toBe("xlsx");
    expect(warnings).toEqual([]);
    expect(workbook.sheets.map((s) => [s.name, s.state ?? "visible"])).toEqual([
      ["Sales", "visible"],
      ["Notes", "visible"],
      ["Hidden", "hidden"],
    ]);
    const sales = workbook.sheets[0]!;
    expect(sales.merges).toEqual(["A1:G1", "A9:C9"]);
    expect(sales.view?.freeze).toMatchObject({ rows: 2 });
    expect(sales.autoFilter).toBe("A2:G8");
    const title = sales.rows[0]![0] as Cell;
    expect(title.value).toBe("Quarterly orders");
    expect(title.style?.font).toMatchObject({ bold: true, size: 16 });
    const revenue = sales.rows[2]![5] as Cell;
    expect(revenue).toMatchObject({ value: 777, formula: "D3*E3" });
    expect(sales.rows[2]![2] && (sales.rows[2]![2] as Cell).value).toBeInstanceOf(Date);
    const link = workbook.sheets[1]!.rows[3]![0] as Cell;
    expect(link.hyperlink).toMatchObject({ target: "https://www.npmjs.com/package/javascript-excel" });
  });

  it("gives the render model what the grid draws", async () => {
    const { workbook } = (await readFile(SAMPLE_XLSX_NAME, await sampleXlsx(), { ...S, read: { ...S.read, dates: "serial" } })).result;
    const view = createRenderSheet(workbook, 0, { format: formatCell });
    expect(sheetSettings(view).freeze).toEqual({ rows: 2, cols: 0 });
    expect(mergeAt(view, 0, 3)?.anchor).toEqual({ row: 0, col: 0 });
    expect(renderCell(view, 0, 0)?.merge).toEqual({ rows: 1, cols: 7 });
    expect(renderCell(view, 2, 5)?.text).toBe("$777.00");
    expect(renderCell(view, 2, 2)?.text).toBe("2026-01-12");
    expect(renderCell(view, 1, 0)?.style.background).toEqual({ kind: "solid", color: "#305496" });
    const notes = createRenderSheet(workbook, 1, { format: formatCell });
    expect(renderCell(notes, 6, 2)).toMatchObject({ text: "(42)", color: "#FF0000" });
    expect(renderCell(notes, 4, 0)?.hyperlink).toMatchObject({ kind: "internal", sheet: 0, ref: "F9" });
  });

  it("honours read options (maxRows, styles: false)", async () => {
    const t = await readFile(SAMPLE_XLSX_NAME, await sampleXlsx(), { ...S, read: { ...S.read, maxRows: 3, styles: false } });
    const sales = t.result.workbook.sheets[0]!;
    expect(sales.rows.length).toBe(3);
    expect((sales.rows[0]![0] as Cell).style).toBeUndefined();
    expect(t.result.warnings.map((w) => w.code)).toContain("ROWS_TRUNCATED");
  });

  it("reads a CSV file with the CSV options", async () => {
    const blob = new Blob([CSV_SAMPLES[1]!.text]);
    const t = await readFile("prices.csv", blob, S);
    expect(t.result.kind).toBe("csv");
    expect(t.result.csv?.delimiter).toBe(";");
    expect(t.result.workbook.sheets[0]!.name).toBe("prices");
    expect(t.result.workbook.sheets[0]!.rows[1]).toEqual(["Äpfel", 12, "1,99", 'enthält "Bio"']);
  });

  it("reports malformed CSV leniently and throws in strict mode", async () => {
    const messy = CSV_SAMPLES.find((c) => c.id === "messy")!.text;
    const lenient = await parseCsvText(messy, S.csvParse, false);
    expect(lenient.result.warnings.map((w) => w.code)).toEqual(expect.arrayContaining(["CSV_STRAY_QUOTE", "CSV_FIELD_COUNT"]));
    await expect(parseCsvText(messy, { ...S.csvParse, mode: "strict" }, false)).rejects.toMatchObject({ code: "CSV_MALFORMED" });
  });

  it("builds the sample dataset with header groups and totals", () => {
    const { sheet } = buildSheet(SAMPLE_DATASET_JSON, S.write).result;
    expect(sheet.merges?.length).toBeGreaterThan(0);
    expect(sheet.rows.map((r) => cellText(r?.[0])).filter(Boolean)).toContain("Total");
    expect(() => buildSheet("{ nope", S.write)).toThrow(/not valid JSON/);
  });
});

describe("display format", () => {
  it("shows Dates without a number format as dates, without changing the model", async () => {
    const { displayFormat } = await import("../src/lib/grid");
    const t = await readFile("d.csv", new Blob(["when\n2026-01-02\n"]), S);
    const wb = t.result.workbook;
    const view = createRenderSheet(wb, 0, { format: displayFormat });
    expect(renderCell(view, 1, 0)?.text).toBe("2026-01-02");
    expect(wb.sheets[0]!.rows[1]![0]).toBeInstanceOf(Date);
  });
});
