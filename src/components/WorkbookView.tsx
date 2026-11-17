// Shows a workbook: sheet tabs, sheet facts, the render-model grid, the cell
// inspector and the warnings. Used by the Read and Write tabs.
import { encodeCell, type CellInput, type ExcelWarning, type Workbook } from "javascript-excel";
import { rangeAsHtml, rangeAsTsv, rangeStats, renderCell, sheetExtent, sheetSettings, type RenderSheet } from "javascript-excel/render";
import { useState } from "react";
import { MAX_PREVIEW_COLS, MAX_PREVIEW_ROWS, rangeLabel, selectionRange, useRenderSheet, type Selection } from "../lib/grid";
import { toJson } from "../lib/json";
import { SheetGrid } from "./SheetGrid";
import { Button } from "./ui";

const START: Selection = { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } };

export function Issues({ warnings }: { warnings: ExcelWarning[] }) {
  if (warnings.length === 0) return <p className="text-sm text-slate-600 dark:text-slate-400">No warnings.</p>;
  return (
    <ul className="max-h-64 space-y-1 overflow-auto text-sm">
      {warnings.map((w, i) => (
        <li key={i} className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100">
          <code className="font-semibold">{w.code}</code>
          {(w.sheet || w.ref) && (
            <span className="ml-1 text-amber-900 dark:text-amber-200">
              {w.sheet ?? ""}
              {w.ref ? `${w.sheet ? "!" : ""}${w.ref}` : ""}
            </span>
          )}
          <span className="ml-1">{w.message}</span>
        </li>
      ))}
    </ul>
  );
}

function Fact({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex gap-1">
      <dt className="text-slate-600 dark:text-slate-400">{label}:</dt>
      <dd className="font-medium text-slate-900 dark:text-slate-100">{children}</dd>
    </div>
  );
}

interface Props {
  workbook: Workbook<CellInput>;
  warnings: ExcelWarning[];
  locale: string;
}

/** Give it a `key` that changes with the workbook, to reset the sheet and selection. */
export function WorkbookView({ workbook, warnings, locale }: Props) {
  const [sheet, setSheet] = useState(workbook.activeSheet ?? 0);
  const [selection, setSelection] = useState<Selection>(START);
  const [copied, setCopied] = useState("");

  const index = workbook.sheets[sheet] ? sheet : 0;
  const view = useRenderSheet(workbook, index, locale);
  const model = workbook.sheets[index];
  if (!model || !view) return <p className="text-sm">The workbook has no sheets.</p>;
  const extent = sheetExtent(view);
  const settings = sheetSettings(view);

  const copy = async (kind: "tsv" | "html"): Promise<void> => {
    const range = selectionRange(selection);
    const text = kind === "tsv" ? rangeAsTsv(view, range) : rangeAsHtml(view, range);
    try {
      await navigator.clipboard.writeText(text);
      setCopied(`Copied ${rangeLabel(range)} as ${kind.toUpperCase()}.`);
    } catch {
      setCopied("The clipboard is not available here.");
    }
  };

  return (
    <div className="space-y-4">
      <div role="tablist" aria-label="Sheets" className="flex flex-wrap gap-1 border-b border-slate-200 dark:border-slate-700">
        {workbook.sheets.map((s, i) => (
          <button
            key={`${s.name}-${i}`}
            role="tab"
            type="button"
            aria-selected={i === index}
            onClick={() => {
              setSheet(i);
              setSelection(START);
              setCopied("");
            }}
            className={`-mb-px rounded-t-md border px-3 py-1.5 text-sm ${i === index ? "border-slate-200 border-b-white bg-white font-semibold text-slate-900 dark:border-slate-700 dark:border-b-slate-900 dark:bg-slate-900 dark:text-white" : "border-transparent text-slate-700 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"}`}
          >
            {s.name}
            {s.state && s.state !== "visible" ? <span className="ml-1 text-xs italic">({s.state})</span> : null}
          </button>
        ))}
      </div>

      <dl className="flex flex-wrap gap-x-5 gap-y-1 text-sm">
        <Fact label="Rows in model">{model.rows.length}</Fact>
        {model.dimension && <Fact label="Dimension">{model.dimension}</Fact>}
        <Fact label="Merges">{model.merges?.length ? model.merges.join(", ") : "none"}</Fact>
        <Fact label="Freeze">{settings.freeze.rows || settings.freeze.cols ? `${settings.freeze.rows} rows, ${settings.freeze.cols} cols` : "none"}</Fact>
        {model.autoFilter && <Fact label="Autofilter">{model.autoFilter}</Fact>}
        {settings.tabColor && (
          <Fact label="Tab colour">
            <span className="inline-flex items-center gap-1">
              <span className="swatch inline-block h-3 w-3 rounded-sm border border-slate-400" style={{ backgroundColor: settings.tabColor }} />
              {settings.tabColor}
            </span>
          </Fact>
        )}
        {settings.rightToLeft && <Fact label="Direction">right to left</Fact>}
      </dl>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-[minmax(0,1fr)_20rem]">
        <div className="min-w-0 space-y-2">
          <SheetGrid view={view} selection={selection} onSelect={(s) => { setSelection(s); setCopied(""); }} label={`Sheet ${model.name}`} />
          {(extent.rows > MAX_PREVIEW_ROWS || extent.cols > MAX_PREVIEW_COLS) && (
            <p className="text-xs text-slate-600 dark:text-slate-400">
              The preview draws the first {Math.min(extent.rows, MAX_PREVIEW_ROWS)} rows and {Math.min(extent.cols, MAX_PREVIEW_COLS)} columns of {extent.rows} × {extent.cols}.
            </p>
          )}
          <div className="flex flex-wrap items-center gap-2">
            <Button onClick={() => void copy("tsv")}>Copy as TSV</Button>
            <Button onClick={() => void copy("html")}>Copy as HTML</Button>
            <span role="status" className="text-sm text-slate-700 dark:text-slate-300">
              {copied}
            </span>
          </div>
        </div>
        <Inspector view={view} workbook={workbook} sheet={index} selection={selection} />
      </div>

      <div>
        <h3 className="mb-2 text-sm font-semibold text-slate-900 dark:text-slate-100">Issues ({warnings.length})</h3>
        <Issues warnings={warnings} />
      </div>

      {workbook.properties && Object.keys(workbook.properties).length > 0 && (
        <details className="text-sm">
          <summary className="cursor-pointer font-semibold text-slate-900 dark:text-slate-100">Document properties</summary>
          <pre className="code mt-2">{toJson(workbook.properties)}</pre>
        </details>
      )}
    </div>
  );
}

function Inspector({ view, workbook, sheet, selection }: { view: RenderSheet; workbook: Workbook<CellInput>; sheet: number; selection: Selection }) {
  const range = selectionRange(selection);
  const { row, col } = selection.focus;
  const cell = renderCell(view, row, col);
  const raw = workbook.sheets[sheet]?.rows[row]?.[col];
  const multi = range.start.row !== range.end.row || range.start.col !== range.end.col;
  const stats = multi ? rangeStats(view, range) : undefined;
  return (
    <aside aria-label="Cell inspector" className="min-w-0 space-y-3 rounded-md border border-slate-200 p-3 text-sm dark:border-slate-700">
      <div className="flex items-baseline justify-between">
        <h3 className="font-semibold text-slate-900 dark:text-slate-100">{multi ? rangeLabel(range) : encodeCell({ row, col })}</h3>
        <span className="text-xs text-slate-600 dark:text-slate-400">{cell?.type ?? "empty"}</span>
      </div>
      {stats && (
        <dl className="grid grid-cols-2 gap-1">
          <Fact label="Count">{stats.count}</Fact>
          <Fact label="Numbers">{stats.numericCount}</Fact>
          <Fact label="Sum">{Number(stats.sum.toFixed(10))}</Fact>
          <Fact label="Average">{stats.average === null ? "–" : Number(stats.average.toFixed(10))}</Fact>
        </dl>
      )}
      <div>
        <div className="text-xs text-slate-600 dark:text-slate-400">Displayed text</div>
        <div className="code min-h-7 whitespace-pre-wrap break-words">{cell?.text || " "}</div>
      </div>
      {cell?.formula && (
        <div>
          <div className="text-xs text-slate-600 dark:text-slate-400">Formula{cell.noCachedValue ? " (no cached value)" : ""}</div>
          <div className="code break-words">{cell.arrayFormula ? `{${cell.formula}}` : cell.formula}</div>
        </div>
      )}
      {cell?.hyperlink && (
        <div>
          <div className="text-xs text-slate-600 dark:text-slate-400">Hyperlink</div>
          <div className="code break-words">
            {cell.hyperlink.kind === "external" ? `${cell.hyperlink.href}${cell.hyperlink.allowed ? "" : ` (blocked: ${cell.hyperlink.reason})`}` : `${cell.hyperlink.ref} on sheet ${cell.hyperlink.sheet === null ? "?" : (workbook.sheets[cell.hyperlink.sheet]?.name ?? "?")}`}
          </div>
        </div>
      )}
      {cell?.merge && <p>Merge anchor: {cell.merge.rows} × {cell.merge.cols}</p>}
      <details open>
        <summary className="cursor-pointer font-medium">Model cell</summary>
        <pre className="code mt-1 max-h-56 overflow-auto">{raw === undefined ? "undefined (no cell)" : toJson(raw)}</pre>
      </details>
      {cell && (
        <details>
          <summary className="cursor-pointer font-medium">Resolved style (render model)</summary>
          <pre className="code mt-1 max-h-56 overflow-auto">{toJson(cell.style)}</pre>
        </details>
      )}
    </aside>
  );
}
