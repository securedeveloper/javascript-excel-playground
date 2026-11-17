// Read: an .xlsx or .csv (the sample or the user's own) into the model.
import type { Progress } from "javascript-excel";
import { useCallback, useEffect, useRef, useState, type Ref } from "react";
import { describeError, formatBytes, formatMs, isCsvName, readFile, type Loaded } from "../lib/engine";
import { SAMPLE_XLSX_NAME, sampleXlsx } from "../lib/samples";
import { LOCALES, type CsvParseSettings, type ReadSettings, type Settings } from "../lib/state";
import { DropZone } from "./DropZone";
import { CsvParseFields } from "./CsvPanel";
import { Button, Card, Fieldset, NumberField, Select, Status, Toggle } from "./ui";
import { WorkbookView } from "./WorkbookView";
import type { RunRecord } from "../App";

export interface Source {
  name: string;
  data: Uint8Array | Blob;
  sample: boolean;
}

interface Props {
  settings: Settings;
  setRead: (r: Partial<ReadSettings>) => void;
  setCsvParse: (c: Partial<CsvParseSettings>) => void;
  onRun: (r: RunRecord) => void;
  fileInputRef: Ref<HTMLInputElement>;
  runSignal: number;
  /** Kept by the parent, so the file survives switching tabs. */
  source: Source | null;
  setSource: (s: Source) => void;
}

export function ReadPanel({ settings, setRead, setCsvParse, onRun, fileInputRef, runSignal, source, setSource }: Props) {
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [status, setStatus] = useState<{ kind: "info" | "error" | "ok"; text: string } | null>(null);
  const [progress, setProgress] = useState<number | null>(null);
  const [busy, setBusy] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const loads = useRef(0);

  const { read: readSettings, csvParse, worker } = settings;

  const run = useCallback(
    async (src: Source) => {
      controller.current?.abort();
      const ac = new AbortController();
      controller.current = ac;
      const id = ++loads.current;
      setBusy(true);
      setProgress(0);
      setStatus({ kind: "info", text: `Reading ${src.name} (${formatBytes(src.data instanceof Blob ? src.data.size : src.data.byteLength)}) ${worker ? "in a Web Worker" : "on the main thread"}…` });
      try {
        const t = await readFile(src.name, src.data, { read: readSettings, csvParse, worker }, { signal: ac.signal, onProgress: (p: Progress) => setProgress(p.ratio) });
        if (id !== loads.current) return;
        setLoaded(t.result);
        const sheets = t.result.workbook.sheets.length;
        const text = `Read ${src.name} in ${formatMs(t.ms)} ${t.where === "worker" ? "(Web Worker)" : "(main thread)"}: ${sheets} sheet${sheets === 1 ? "" : "s"}, ${t.result.warnings.length} warning${t.result.warnings.length === 1 ? "" : "s"}.`;
        setStatus({ kind: "ok", text });
        onRun({ source: src.sample ? "sample" : src.name, action: t.result.kind === "csv" ? "readCsv" : "read", ms: t.ms, where: t.where, summary: `${sheets} sheet(s), ${t.result.warnings.length} warning(s)`, ok: true });
      } catch (e) {
        if (id !== loads.current) return;
        setLoaded(null);
        setStatus({ kind: "error", text: describeError(e) });
        onRun({ source: src.sample ? "sample" : src.name, action: isCsvName(src.name) ? "readCsv" : "read", ms: 0, where: worker ? "worker" : "main", summary: describeError(e), ok: false });
      } finally {
        if (id === loads.current) {
          setBusy(false);
          setProgress(null);
          controller.current = null;
        }
      }
    },
    [readSettings, csvParse, worker, onRun],
  );

  const loadSample = useCallback(async () => {
    const bytes = await sampleXlsx();
    setSource({ name: SAMPLE_XLSX_NAME, data: bytes, sample: true });
  }, [setSource]);

  // Start with the sample, so there is something to look at.
  const hasSource = source !== null;
  useEffect(() => {
    if (!hasSource) void loadSample();
  }, [hasSource, loadSample]);

  // Re-read the current file whenever the settings change, or on Ctrl+Enter.
  useEffect(() => {
    // Starts an async read (like a fetch); its busy and progress state is set as it begins.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    if (source) void run(source);
  }, [source, run, runSignal]);

  const isCsv = source ? isCsvName(source.name) : false;

  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-[17rem_minmax(0,1fr)]">
      <div className="space-y-4">
        <Card title="Source">
          <DropZone accept=".xlsx,.xlsm,.xltx,.xltm,.csv,.tsv,.txt" onFile={(f) => setSource({ name: f.name, data: f, sample: false })} inputRef={fileInputRef} label="Open a file" hint="or drop an .xlsx or .csv here. It stays in your browser; nothing is uploaded." />
          <div className="mt-3 flex flex-wrap gap-2">
            <Button onClick={() => void loadSample()}>Load the sample workbook</Button>
            {busy && (
              <Button onClick={() => controller.current?.abort()} variant="ghost">
                Cancel
              </Button>
            )}
          </div>
          {source && (
            <p className="mt-2 break-all text-xs text-slate-600 dark:text-slate-400">
              Current: {source.name}
              {source.sample ? " (generated with the library)" : ""}
            </p>
          )}
        </Card>
        <Card title="Read options">
          <div className="space-y-3">
            <Fieldset legend=".xlsx">
              <Toggle label="Styles" hint="styles: false is faster; no cell.style" checked={readSettings.styles} onChange={(v) => setRead({ styles: v })} />
              <Toggle label="Formulas" hint="formulas: false keeps cached values only" checked={readSettings.formulas} onChange={(v) => setRead({ formulas: v })} />
              <Select label="Dates" value={readSettings.dates} options={[["date", "date: Date objects"], ["serial", "serial: numbers (exact)"]]} onChange={(v) => setRead({ dates: v })} />
              <Select label="Rich text" value={readSettings.richText} options={[["runs", "runs"], ["plain", "plain"]]} onChange={(v) => setRead({ richText: v })} />
              <NumberField label="Max rows" hint="0 = all rows" value={readSettings.maxRows} min={0} max={1_048_576} onChange={(v) => setRead({ maxRows: v })} />
              <Select label="Display locale" value={readSettings.locale} options={LOCALES.map((l) => [l, l] as const)} onChange={(v) => setRead({ locale: v })} />
            </Fieldset>
            <Fieldset legend=".csv">
              <CsvParseFields value={csvParse} onChange={setCsvParse} withEncoding />
            </Fieldset>
          </div>
        </Card>
      </div>
      <Card title={loaded ? `Result: ${loaded.name}` : "Result"}>
        <div className="space-y-3">
          {progress !== null && <progress className="h-2 w-full" max={1} value={progress} aria-label="Read progress" />}
          {status && <Status kind={status.kind}>{status.text}</Status>}
          {loaded?.csv && (
            <p className="text-sm text-slate-700 dark:text-slate-300">
              CSV: delimiter <code>{loaded.csv.delimiter === "\t" ? "TAB" : loaded.csv.delimiter}</code>, newline <code>{JSON.stringify(loaded.csv.newline)}</code>
              {loaded.csv.bom ? ", with BOM" : ""}.
            </p>
          )}
          {loaded && <WorkbookView key={`${loaded.name}:${loaded.size}`} workbook={loaded.workbook} warnings={loaded.warnings} locale={readSettings.locale} />}
          {!loaded && !status && <p className="text-sm">Loading the sample…</p>}
          {isCsv && <p className="text-xs text-slate-600 dark:text-slate-400">CSV files are read with the .csv options; the .xlsx options don't apply.</p>}
        </div>
      </Card>
    </div>
  );
}
