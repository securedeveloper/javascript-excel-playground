// Built-in samples. The workbook is generated with the library itself, so the
// repository ships no binary fixtures and the sample always matches the
// published version.
import { toUint8Array, type Cell, type Style, type Workbook } from "javascript-excel";

const border = { style: "thin", color: "A6A6A6" } as const;
const boxed = { top: border, bottom: border, left: border, right: border };
const title: Style = { font: { bold: true, size: 16, color: "1F3864" }, alignment: { horizontal: "center", vertical: "center" } };
const header: Style = { font: { bold: true, color: "FFFFFF" }, fill: "305496", border: boxed, alignment: { horizontal: "center" } };
const text: Style = { border: boxed };
const date: Style = { numFmt: "yyyy-mm-dd", border: boxed };
const money: Style = { numFmt: '"$"#,##0.00', border: boxed };
const pct: Style = { numFmt: "0.0%", border: boxed, font: { color: "375623" } };
const total: Style = { font: { bold: true }, fill: "E2EFDA", border: { ...boxed, top: { style: "double", color: "375623" } }, numFmt: '"$"#,##0.00' };

const ORDERS: Array<[string, string, string, number, number]> = [
  ["North", "Ada", "2026-01-12", 42, 18.5],
  ["North", "Grace", "2026-02-03", 17, 22],
  ["South", "Linus", "2026-02-19", 63, 9.75],
  ["East", "Margaret", "2026-03-07", 28, 31.2],
  ["West", "Alan", "2026-03-22", 51, 14],
  ["South", "Barbara", "2026-04-02", 9, 120],
];

/** A small workbook that shows what the reader returns: styles, number formats, merges, formulas, links, rich text. */
export function sampleWorkbook(): Workbook {
  const rows: Cell[][] = [
    [{ value: "Quarterly orders", style: title }],
    ["Region", "Rep", "Date", "Units", "Unit price", "Revenue", "Share"].map((v) => ({ value: v, style: header })),
  ];
  const grand = ORDERS.reduce((s, o) => s + o[3] * o[4], 0);
  ORDERS.forEach(([region, rep, iso, units, price], i) => {
    const r = i + 3;
    const [y, m, d] = iso.split("-").map(Number) as [number, number, number];
    const revenue = units * price;
    rows.push([
      { value: region, style: text },
      { value: rep, style: text },
      { value: new Date(y, m - 1, d), style: date },
      { value: units, style: { ...text, numFmt: "#,##0" } },
      { value: price, style: money },
      { value: revenue, formula: `D${r}*E${r}`, style: money },
      { value: revenue / grand, formula: `F${r}/$F$${ORDERS.length + 3}`, style: pct },
    ]);
  });
  const last = ORDERS.length + 2;
  rows.push([
    { value: "Total", style: { ...total, numFmt: undefined } },
    { value: null, style: total },
    { value: null, style: total },
    { value: ORDERS.reduce((s, o) => s + o[3], 0), formula: `SUM(D3:D${last})`, style: { ...total, numFmt: "#,##0" } },
    { value: null, style: total },
    { value: grand, formula: `SUM(F3:F${last})`, style: total },
    { value: 1, formula: `SUM(G3:G${last})`, style: { ...total, numFmt: "0.0%" } },
  ]);

  const notes: Cell[][] = [
    [{ value: "Notes", style: { font: { bold: true, size: 14 } } }],
    [
      {
        value: {
          richText: [
            { text: "Rich text: " },
            { text: "bold", font: { bold: true } },
            { text: ", " },
            { text: "italic red", font: { italic: true, color: "C00000" } },
          ],
        },
      },
    ],
    [{ value: "Wrapped text keeps\nits line breaks.", style: { alignment: { wrapText: true, vertical: "top" } } }],
    [{ value: "javascript-excel on npm", hyperlink: { target: "https://www.npmjs.com/package/javascript-excel", tooltip: "Opens npm" }, style: { font: { color: "0563C1", underline: true } } }],
    [{ value: "Jump to Sales!F9", hyperlink: { location: "'Sales'!F9" }, style: { font: { color: "0563C1", underline: true } } }],
    [{ value: true }, { value: false }, { value: { error: "#N/A" } }],
    [{ value: 0.5, style: { numFmt: "# ?/?" } }, { value: 1234567.891, style: { numFmt: "#,##0.00,, \"M\"" } }, { value: -42, style: { numFmt: "#,##0;[Red](#,##0)" } }],
  ];

  return {
    properties: {
      title: "javascript-excel playground sample",
      creator: "javascript-excel playground",
      created: new Date(Date.UTC(2026, 0, 1)),
      modified: new Date(Date.UTC(2026, 0, 1)),
    },
    sheets: [
      {
        name: "Sales",
        rows,
        merges: ["A1:G1", `A${last + 1}:C${last + 1}`],
        rowProps: [{ height: 28 }],
        colProps: [{ width: 10 }, { width: 12 }, { width: 12 }, { width: 8 }, { width: 12 }, { width: 13 }, { width: 9 }],
        view: { freeze: { rows: 2 } },
        autoFilter: `A2:G${last}`,
        tabColor: "305496",
      },
      { name: "Notes", rows: notes, colProps: [{ width: 28 }, { width: 14 }, { width: 12 }], rowProps: [undefined, undefined, { height: 32 }] },
      { name: "Hidden", state: "hidden", rows: [["This sheet is hidden in Excel."]] },
    ],
  };
}

export const SAMPLE_XLSX_NAME = "sample-orders.xlsx";

/** The sample workbook as .xlsx bytes. */
export function sampleXlsx(): Promise<Uint8Array> {
  return toUint8Array(sampleWorkbook());
}

export interface CsvSample {
  id: string;
  label: string;
  text: string;
}

export const CSV_SAMPLES: readonly CsvSample[] = [
  {
    id: "orders",
    label: "Orders (comma, typed values)",
    text: ["region,rep,date,units,price,active", "North,Ada,2026-01-12,42,18.5,TRUE", "South,Linus,2026-02-19,63,9.75,FALSE", '"West, coast",Alan,2026-03-22,51,14,TRUE', "East,Margaret,2026-03-07,007,31.2,true"].join("\r\n"),
  },
  {
    id: "semicolon",
    label: "European (semicolon, quoted)",
    text: ['Produkt;Menge;Preis;Notiz', 'Äpfel;12;"1,99";"enthält ""Bio"""', 'Birnen;7;"2,49";', 'Kirschen;30;"5,00";"mehrzeilige', 'Notiz"'].join("\n"),
  },
  {
    id: "messy",
    label: "Messy (stray quote, ragged rows)",
    text: ['id,name,comment', '1,Ada,fine', '2,Gr"ace,stray quote', '3,Linus', '4,Alan,"extra",field', '', '5,"=HYPERLINK(""x"")",formula-looking text'].join("\n"),
  },
];

export const SAMPLE_DATASET = {
  columns: [
    { title: "Name", key: "name" },
    { title: "Joined", key: "joined", type: "date", numFmt: "yyyy-mm-dd" },
    {
      title: "Sales",
      children: [
        { title: "Q1", path: "sales.q1", numFmt: "#,##0.00", total: "sum" },
        { title: "Q2", path: "sales.q2", numFmt: "#,##0.00", total: "sum" },
      ],
    },
    { title: "Active", key: "active" },
  ],
  data: [
    { name: "Ada", joined: "2024-01-15", sales: { q1: 1200.5, q2: 980 }, active: true },
    { name: "Grace", joined: "2023-06-01", sales: { q1: 860, q2: 1430.25 }, active: true },
    { name: "Linus", joined: "2022-11-30", sales: { q1: 0, q2: 75.5 }, active: false },
  ],
};

export const SAMPLE_DATASET_JSON = JSON.stringify(SAMPLE_DATASET, null, 2);
