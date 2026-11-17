// CSV: parse text with the chosen options, inspect the typed rows, write it back.
import type { CellValue } from "javascript-excel";
import { CSV_MIME, type CsvParseResult } from "javascript-excel/csv";
import { download } from "javascript-excel/dom";
import { useEffect, useId, useState } from "react";
import { csvTextOf, describeError, formatMs, parseCsvText } from "../lib/engine";
import { CSV_SAMPLES } from "../lib/samples";
import type { CsvParseSettings, CsvWriteSettings, Settings } from "../lib/state";
import { Button, Card, Fieldset, Select, Status, Toggle } from "./ui";
import { Issues } from "./WorkbookView";
import type { RunRecord } from "../App";

const DELIMS = [
  ["auto", "auto-detect"],
  [",", "comma ,"],
  [";", "semicolon ;"],
  ["\t", "tab"],
  ["|", "pipe |"],
] as const;

export function CsvParseFields({ value, onChange, withEncoding = false }: { value: CsvParseSettings; onChange: (v: Partial<CsvParseSettings>) => void; withEncoding?: boolean }) {
  return (
    <>
      <Select label="Delimiter" value={value.delimiter} options={DELIMS} onChange={(v) => onChange({ delimiter: v })} />
      <Toggle label="Infer types" hint="numbers, TRUE/FALSE and ISO dates; quoted fields stay text" checked={value.inferTypes} onChange={(v) => onChange({ inferTypes: v })} />
      <Toggle label="Skip empty lines" checked={value.skipEmptyLines} onChange={(v) => onChange({ skipEmptyLines: v })} />
      <Select label="Mode" value={value.mode} options={[["lenient", "lenient: warnings"], ["strict", "strict: throw CSV_MALFORMED"]]} onChange={(v) => onChange({ mode: v })} />
      {withEncoding && (
        <Select
          label="Encoding (without a BOM)"
          value={value.encoding}
          options={[
            ["utf-8", "UTF-8"],
            ["utf-16le", "UTF-16LE"],
            ["windows-1252", "Windows-1252"],
            ["iso-8859-1", "ISO-8859-1"],
          ]}
          onChange={(v) => onChange({ encoding: v })}
        />
      )}
    </>
  );
}

function typeOf(v: CellValue | undefined): string {
  if (v === undefined || v === "") return "empty";
  if (v instanceof Date) return "date";
  if (typeof v === "object") return "error" in v ? "error" : "richText";
  return typeof v;
}

function show(v: CellValue | undefined): string {
  if (v instanceof Date) {
    // Local wall-clock time, as the CSV wrote it (zone-less ISO dates are read as local).
    const p = (n: number): string => String(n).padStart(2, "0");
    const date = `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
    const time = v.getHours() || v.getMinutes() || v.getSeconds() ? ` ${p(v.getHours())}:${p(v.getMinutes())}:${p(v.getSeconds())}` : "";
    return date + time;
  }
  if (v && typeof v === "object") return "error" in v ? v.error : v.richText.map((r) => r.text).join("");
  return v === undefined ? "" : String(v);
}

const TYPE_CLASS: Record<string, string> = {
  number: "text-blue-800 dark:text-blue-300 text-right",
  boolean: "text-purple-800 dark:text-purple-300 text-center",
  date: "text-green-800 dark:text-green-300",
};

interface Props {
  settings: Settings;
  setCsvText: (t: string) => void;
  setCsvParse: (c: Partial<CsvParseSettings>) => void;
  setCsvWrite: (c: Partial<CsvWriteSettings>) => void;
  onRun: (r: RunRecord) => void;
  runSignal: number;
}

export function CsvPanel({ settings, setCsvText, setCsvParse, setCsvWrite, onRun, runSignal }: Props) {
  const { csvText, csvParse, csvWrite, worker } = settings;
  const [parsed, setParsed] = useState<{ result: CsvParseResult; ms: number; where: "main" | "worker" } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [outputState, setOutput] = useState<{ text: string; ms: number; for: unknown } | null>(null);
  const areaId = useId();

  useEffect(() => {
    let live = true;
    const t = setTimeout(() => {
      parseCsvText(csvText, csvParse, worker).then(
        (r) => {
          if (!live) return;
          setParsed(r);
          setError(null);
          onRun({ source: "CSV text", action: "parseCsv", ms: r.ms, where: r.where, summary: `${r.result.rows.length} rows, ${r.result.warnings.length} warning(s)`, ok: true });
        },
        (e: unknown) => {
          if (!live) return;
          setParsed(null);
          setError(describeError(e));
          onRun({ source: "CSV text", action: "parseCsv", ms: 0, where: worker ? "worker" : "main", summary: describeError(e), ok: false });
        },
      );
    }, 250);
    return () => {
      live = false;
      clearTimeout(t);
    };
  }, [csvText, csvParse, worker, onRun, runSignal]);

  useEffect(() => {
    let live = true;
    if (!parsed) return;
    csvTextOf(parsed.result.rows, csvWrite, worker).then(
      (r) => live && setOutput({ text: r.result, ms: r.ms, for: parsed }),
      (e: unknown) => live && setOutput({ text: describeError(e), ms: 0, for: parsed }),
    );
    return () => {
      live = false;
    };
  }, [parsed, csvWrite, worker]);

  const output = parsed && outputState?.for === parsed ? outputState : null;
  const rows = parsed?.result.rows ?? [];
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0);

  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
      <Card
        title="CSV input"
        actions={
          <Select
            label="Sample"
            value={CSV_SAMPLES.find((s) => s.text === csvText)?.id ?? ""}
            options={[["", "Custom text"], ...CSV_SAMPLES.map((s) => [s.id, s.label] as const)]}
            onChange={(id) => {
              const s = CSV_SAMPLES.find((x) => x.id === id);
              if (s) setCsvText(s.text);
            }}
          />
        }
      >
        <label htmlFor={areaId} className="sr-only">
          CSV text
        </label>
        <textarea id={areaId} value={csvText} onChange={(e) => setCsvText(e.target.value)} spellCheck={false} rows={10} className="field code-input w-full" />
        <div className="mt-3">
          <Fieldset legend="Parse options (parseCsv)">
            <div className="grid gap-x-4 sm:grid-cols-2">
              <CsvParseFields value={csvParse} onChange={setCsvParse} />
            </div>
          </Fieldset>
        </div>
      </Card>

      <Card title="Parsed rows">
        <div className="space-y-3">
          {error && <Status kind="error">{error}</Status>}
          {parsed && (
            <Status kind="ok">
              Parsed {rows.length} rows in {formatMs(parsed.ms)} ({parsed.where === "worker" ? "Web Worker" : "main thread"}). Delimiter <code>{parsed.result.delimiter === "\t" ? "TAB" : parsed.result.delimiter}</code>, newline <code>{JSON.stringify(parsed.result.newline)}</code>
              {parsed.result.bom ? ", BOM" : ""}.
            </Status>
          )}
          {parsed && (
            <div className="max-h-80 overflow-auto rounded-md border border-slate-200 dark:border-slate-700" tabIndex={0} role="region" aria-label="Parsed rows">
              <table className="csv-table w-full text-sm">
                <thead>
                  <tr>
                    <th scope="col">#</th>
                    {Array.from({ length: width }, (_, c) => (
                      <th key={c} scope="col">
                        {c + 1}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 200).map((r, i) => (
                    <tr key={i}>
                      <th scope="row">{i + 1}</th>
                      {Array.from({ length: width }, (_, c) => {
                        const v = r[c];
                        const t = typeOf(v);
                        return (
                          <td key={c} className={TYPE_CLASS[t] ?? ""} title={t}>
                            {show(v)}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          <p className="text-xs text-slate-600 dark:text-slate-400">
            Colours: <span className="text-blue-800 dark:text-blue-300">number</span>, <span className="text-purple-800 dark:text-purple-300">boolean</span>, <span className="text-green-800 dark:text-green-300">date</span>, text. Hover a cell for its type.
          </p>
          <div>
            <h3 className="mb-2 text-sm font-semibold">Issues ({parsed?.result.warnings.length ?? 0})</h3>
            <Issues warnings={parsed?.result.warnings ?? []} />
          </div>
        </div>
      </Card>

      <Card
        title="Write it back (toCsv)"
        className="lg:col-span-2"
        actions={
          <Button
            variant="primary"
            disabled={!output || !parsed}
            onClick={() => {
              if (output) download(output.text, "playground.csv", { mimeType: CSV_MIME });
            }}
          >
            Download .csv
          </Button>
        }
      >
        <div className="grid grid-cols-1 gap-4 md:grid-cols-[16rem_minmax(0,1fr)]">
          <Fieldset legend="Write options">
            <Select label="Delimiter" value={csvWrite.delimiter} options={DELIMS.slice(1) as ReadonlyArray<readonly [CsvWriteSettings["delimiter"], string]>} onChange={(v) => setCsvWrite({ delimiter: v })} />
            <Select label="Newline" value={csvWrite.newline} options={[["\r\n", "CRLF (RFC 4180)"], ["\n", "LF"]]} onChange={(v) => setCsvWrite({ newline: v })} />
            <Select label="Quote" value={csvWrite.quote} options={[["auto", "auto"], ["all", "all fields"]]} onChange={(v) => setCsvWrite({ quote: v })} />
            <Toggle label="UTF-8 BOM" hint="so Excel detects UTF-8" checked={csvWrite.bom} onChange={(v) => setCsvWrite({ bom: v })} />
            <Toggle label="Formula-injection guard" hint="escapeFormulas: prefix = + - @ with '" checked={csvWrite.escapeFormulas} onChange={(v) => setCsvWrite({ escapeFormulas: v })} />
          </Fieldset>
          <div className="min-w-0">
            {output && <p className="mb-2 text-sm text-slate-700 dark:text-slate-300">Serialised in {formatMs(output.ms)}{output.text.startsWith("\uFEFF") ? "; it starts with a BOM (not shown)" : ""}. Line ends are shown as ⏎.</p>}
            <pre className="code max-h-72 overflow-auto whitespace-pre" aria-label="CSV output">
              {output?.text.replace(/^\uFEFF/, "").replace(/\r?\n/g, (m) => `⏎${m}`) ?? ""}
            </pre>
          </div>
        </div>
      </Card>
    </div>
  );
}
