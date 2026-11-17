// Write: a dataset ({ columns, data }) to a sheet, previewed, then .xlsx or .csv.
import type { ExcelWarning, Workbook } from "javascript-excel";
import { toCsvBlob } from "javascript-excel/csv";
import { download } from "javascript-excel/dom";
import { read } from "javascript-excel/read";
import { useEffect, useId, useMemo, useState } from "react";
import { buildSheet, csvWriteOptions, describeError, formatBytes, formatMs, writeXlsx } from "../lib/engine";
import { SAMPLE_DATASET_JSON } from "../lib/samples";
import type { Settings, WriteSettings } from "../lib/state";
import { Button, Card, Fieldset, Select, Status, TextField, Toggle } from "./ui";
import { WorkbookView } from "./WorkbookView";
import type { RunRecord } from "../App";

interface Props {
  settings: Settings;
  setWrite: (w: Partial<WriteSettings>) => void;
  setDataset: (json: string) => void;
  onRun: (r: RunRecord) => void;
  runSignal: number;
}

export function WritePanel({ settings, setWrite, setDataset, onRun, runSignal }: Props) {
  const { write, dataset, worker } = settings;
  // Status and round trip belong to the sheet they were made from; a rebuilt sheet hides them.
  const [statusState, setStatusState] = useState<{ kind: "info" | "error" | "ok"; text: string; for: unknown } | null>(null);
  const [roundTripState, setRoundTrip] = useState<{ workbook: Workbook; warnings: ExcelWarning[]; key: string; for: unknown } | null>(null);
  const areaId = useId();

  const built = useMemo(() => {
    try {
      return { ok: true as const, ...buildSheet(dataset, write) };
    } catch (e) {
      return { ok: false as const, error: describeError(e) };
    }
    // runSignal re-builds on Ctrl+Enter
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dataset, write, runSignal]);

  useEffect(() => {
    if (built.ok) onRun({ source: "dataset", action: "fromDataset", ms: built.ms, where: "main", summary: `${built.result.sheet.rows.length} rows`, ok: true });
    else onRun({ source: "dataset", action: "fromDataset", ms: 0, where: "main", summary: built.error, ok: false });
  }, [built, onRun]);

  const workbook = useMemo<Workbook | null>(() => (built.ok ? { sheets: [built.result.sheet], properties: { title: "javascript-excel playground" } } : null), [built]);
  const status = statusState?.for === built ? statusState : null;
  const roundTrip = roundTripState?.for === built ? roundTripState : null;
  const setStatus = (s: { kind: "info" | "error" | "ok"; text: string }): void => setStatusState({ ...s, for: built });

  const saveXlsx = async (): Promise<void> => {
    if (!workbook) return;
    try {
      const t = await writeXlsx(workbook, write, worker);
      download(t.result, "dataset.xlsx");
      setStatus({ kind: "ok", text: `Wrote dataset.xlsx (${formatBytes(t.result.byteLength)}) in ${formatMs(t.ms)} ${t.where === "worker" ? "(Web Worker)" : "(main thread)"}.` });
      onRun({ source: "dataset", action: "toUint8Array", ms: t.ms, where: t.where, summary: formatBytes(t.result.byteLength), ok: true });
    } catch (e) {
      setStatus({ kind: "error", text: describeError(e) });
    }
  };

  const saveCsv = (): void => {
    if (!built.ok) return;
    try {
      const started = performance.now();
      const blob = toCsvBlob(built.result.sheet, csvWriteOptions(settings.csvWrite));
      const ms = performance.now() - started;
      download(blob, "dataset.csv");
      setStatus({ kind: "ok", text: `Wrote dataset.csv (${formatBytes(blob.size)}) in ${formatMs(ms)}, with the CSV tab's write options.` });
      onRun({ source: "dataset", action: "toCsvBlob", ms, where: "main", summary: formatBytes(blob.size), ok: true });
    } catch (e) {
      setStatus({ kind: "error", text: describeError(e) });
    }
  };

  const readBack = async (): Promise<void> => {
    if (!workbook) return;
    try {
      const w = await writeXlsx(workbook, write, worker);
      const started = performance.now();
      const r = await read(w.result);
      const ms = performance.now() - started;
      setRoundTrip({ workbook: r.workbook, warnings: r.warnings, key: String(Date.now()), for: built });
      setStatus({ kind: "ok", text: `Wrote ${formatBytes(w.result.byteLength)} in ${formatMs(w.ms)}, read it back in ${formatMs(ms)}. Below is what the reader returns.` });
      onRun({ source: "dataset", action: "write + read", ms: w.ms + ms, where: w.where, summary: `${formatBytes(w.result.byteLength)} round trip`, ok: true });
    } catch (e) {
      setStatus({ kind: "error", text: describeError(e) });
    }
  };

  return (
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-[21rem_minmax(0,1fr)]">
      <div className="space-y-4">
        <Card title="Dataset (JSON)" actions={<Button onClick={() => setDataset(SAMPLE_DATASET_JSON)}>Reset</Button>}>
          <label htmlFor={areaId} className="sr-only">
            Dataset JSON
          </label>
          <textarea id={areaId} value={dataset} onChange={(e) => setDataset(e.target.value)} spellCheck={false} rows={16} className="field code-input w-full" aria-describedby={`${areaId}-help`} />
          <p id={`${areaId}-help`} className="mt-2 text-xs text-slate-600 dark:text-slate-400">
            <code>{"{ columns, data }"}</code>. Columns take key, path, title, type, numFmt, width, total and children (header groups).
          </p>
        </Card>
        <Card title="Options (fromDataset)">
          <TextField label="Sheet name" value={write.sheetName} maxLength={31} onChange={(v) => setWrite({ sheetName: v })} />
          <Toggle label="Freeze header rows" checked={write.freezeHeader} onChange={(v) => setWrite({ freezeHeader: v })} />
          <Toggle label="Autofilter" checked={write.autoFilter} onChange={(v) => setWrite({ autoFilter: v })} />
          <Toggle label="Auto column width" checked={write.autoWidth} onChange={(v) => setWrite({ autoWidth: v })} />
          <Toggle label="Totals row label" hint="columns with total: 'sum' get SUM formulas" checked={write.totals} onChange={(v) => setWrite({ totals: v })} />
          <Toggle label="Banded rows" hint="rowStyle: a fill on every other row" checked={write.banded} onChange={(v) => setWrite({ banded: v })} />
          <Select
            label="Header fill"
            value={write.headerFill}
            options={[
              ["", "none"],
              ["DDEBF7", "light blue"],
              ["E2EFDA", "light green"],
              ["FCE4D6", "light orange"],
            ]}
            onChange={(v) => setWrite({ headerFill: v })}
          />
          <Fieldset legend="Writer (toUint8Array)">
            <Select label="Compression" value={write.compression} options={[["deflate", "deflate (smaller)"], ["store", "store (faster)"]]} onChange={(v) => setWrite({ compression: v })} />
          </Fieldset>
        </Card>
      </div>
      <Card
        title="Preview"
        actions={
          <>
            <Button variant="primary" disabled={!workbook} onClick={() => void saveXlsx()}>
              Download .xlsx
            </Button>
            <Button disabled={!workbook} onClick={saveCsv}>
              Download .csv
            </Button>
            <Button disabled={!workbook} onClick={() => void readBack()}>
              Write and read back
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          {!built.ok && <Status kind="error">{built.error}</Status>}
          {built.ok && !status && <Status kind="info">Built a sheet of {built.result.sheet.rows.length} rows in {formatMs(built.ms)}. The preview is the render model of the sheet before it is written.</Status>}
          {status && <Status kind={status.kind}>{status.text}</Status>}
          {roundTrip ? (
            <>
              <WorkbookView key={roundTrip.key} workbook={roundTrip.workbook} warnings={roundTrip.warnings} locale={settings.read.locale} />
              <Button variant="ghost" onClick={() => setRoundTrip(null)}>
                Back to the built sheet
              </Button>
            </>
          ) : (
            workbook && built.ok && <WorkbookView key={`${dataset.length}:${write.sheetName}`} workbook={workbook} warnings={built.result.warnings} locale={settings.read.locale} />
          )}
        </div>
      </Card>
    </div>
  );
}
