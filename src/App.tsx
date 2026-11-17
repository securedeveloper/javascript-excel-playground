import { VERSION } from "javascript-excel";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { CodePanel } from "./components/CodePanel";
import { CsvPanel } from "./components/CsvPanel";
import { HistoryPanel } from "./components/HistoryPanel";
import { ReadPanel, type Source } from "./components/ReadPanel";
import { ShortcutsDialog, type Shortcut } from "./components/ShortcutsDialog";
import { Button, Card } from "./components/ui";
import { WritePanel } from "./components/WritePanel";
import { useTheme } from "./hooks/useTheme";
import { snippetFor } from "./lib/codegen";
import { isCsvName } from "./lib/engine";
import { EXAMPLES } from "./lib/examples";
import { addEntry, loadHistory, saveHistory, type HistoryEntry } from "./lib/history";
import { CSV_SAMPLES, SAMPLE_DATASET_JSON } from "./lib/samples";
import { decodeSettings, defaultSettings, encodeSettings, TABS, type CsvParseSettings, type CsvWriteSettings, type ReadSettings, type Settings, type Tab, type WriteSettings } from "./lib/state";

export type RunRecord = Omit<HistoryEntry, "id" | "at" | "tab" | "settings">;

const DEFAULTS = defaultSettings(SAMPLE_DATASET_JSON, CSV_SAMPLES[0]!.text);

const LINKS = {
  github: "https://github.com/securedeveloper/javascript-excel",
  npm: "https://www.npmjs.com/package/javascript-excel",
  docs: "https://github.com/securedeveloper/javascript-excel/blob/main/docs/api.md",
  playground: "https://github.com/securedeveloper/javascript-excel-playground",
};

const TAB_LABELS: Record<Tab, string> = { read: "Read", write: "Write", csv: "CSV" };

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);
const MOD = isMac ? "⌘" : "Ctrl";

const SHORTCUTS: Shortcut[] = [
  { keys: `${MOD} + Enter`, description: "Run again with the current settings" },
  { keys: "Alt + 1 / 2 / 3", description: "Switch to Read, Write or CSV" },
  { keys: `${MOD} + O`, description: "Open a file (Read tab)" },
  { keys: "Alt + W", description: "Toggle Web Worker mode" },
  { keys: "Alt + T", description: "Toggle light and dark theme" },
  { keys: "Alt + S", description: "Copy a link to the current settings" },
  { keys: "Arrows, Shift, Ctrl/⌘", description: "In a sheet: move, extend the selection, jump to the data edge" },
  { keys: "?", description: "Show this help" },
  { keys: "Esc", description: "Close this help" },
];

function isTyping(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  return !!el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT" || el.isContentEditable);
}

export default function App() {
  const [settings, setSettings] = useState<Settings>(() => (typeof location === "undefined" ? DEFAULTS : decodeSettings(location.search, DEFAULTS)));
  const [runs, setRuns] = useState<HistoryEntry[]>(loadHistory);
  const [runSignal, setRunSignal] = useState(0);
  const [help, setHelp] = useState(false);
  const [shared, setShared] = useState("");
  const [source, setSource] = useState<Source | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const theme = useTheme();
  const settingsRef = useRef(settings);
  useEffect(() => {
    settingsRef.current = settings;
  }, [settings]);

  // Keep the URL in step with the settings, so a link reproduces them.
  useEffect(() => {
    const q = encodeSettings(settings, DEFAULTS);
    const url = `${location.pathname}${q ? `?${q}` : ""}${location.hash}`;
    if (url !== `${location.pathname}${location.search}${location.hash}`) window.history.replaceState(null, "", url);
  }, [settings]);

  useEffect(() => saveHistory(runs), [runs]);

  const onRun = useCallback((r: RunRecord) => {
    const { dataset: _d, csvText: _c, ...rest } = settingsRef.current;
    void _d;
    void _c;
    setRuns((h) => addEntry(h, { ...r, tab: rest.tab, settings: rest }));
  }, []);

  const setTab = useCallback((tab: Tab) => setSettings((s) => ({ ...s, tab })), []);
  const setRead = useCallback((r: Partial<ReadSettings>) => setSettings((s) => ({ ...s, read: { ...s.read, ...r } })), []);
  const setCsvParse = useCallback((c: Partial<CsvParseSettings>) => setSettings((s) => ({ ...s, csvParse: { ...s.csvParse, ...c } })), []);
  const setCsvWrite = useCallback((c: Partial<CsvWriteSettings>) => setSettings((s) => ({ ...s, csvWrite: { ...s.csvWrite, ...c } })), []);
  const setWrite = useCallback((w: Partial<WriteSettings>) => setSettings((s) => ({ ...s, write: { ...s.write, ...w } })), []);
  const setDataset = useCallback((dataset: string) => setSettings((s) => ({ ...s, dataset })), []);
  const setCsvText = useCallback((csvText: string) => setSettings((s) => ({ ...s, csvText })), []);

  const share = useCallback(() => {
    const q = encodeSettings(settingsRef.current, DEFAULTS);
    const url = `${location.origin}${location.pathname}${q ? `?${q}` : ""}`;
    navigator.clipboard.writeText(url).then(
      () => setShared("Link copied."),
      () => setShared("Could not copy; the address bar has the link."),
    );
    setTimeout(() => setShared(""), 2500);
  }, []);

  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      const mod = isMac ? e.metaKey : e.ctrlKey;
      if (mod && e.key === "Enter") {
        e.preventDefault();
        setRunSignal((n) => n + 1);
      } else if (mod && e.key.toLowerCase() === "o") {
        e.preventDefault();
        setTab("read");
        setTimeout(() => fileInput.current?.click(), 0);
      } else if (e.altKey && !mod && /^Digit[1-3]$/.test(e.code)) {
        e.preventDefault();
        setTab(TABS[Number(e.code.slice(5)) - 1]!);
      } else if (e.altKey && e.code === "KeyW") {
        e.preventDefault();
        setSettings((s) => ({ ...s, worker: !s.worker }));
      } else if (e.altKey && e.code === "KeyT") {
        e.preventDefault();
        theme.toggle();
      } else if (e.altKey && e.code === "KeyS") {
        e.preventDefault();
        share();
      } else if (e.key === "?" && !isTyping(e.target)) {
        e.preventDefault();
        setHelp(true);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [setTab, share, theme]);

  const fileKind = source && isCsvName(source.name) ? "csv" : "xlsx";
  const snippet = useMemo(() => snippetFor(settings, fileKind), [settings, fileKind]);

  const restore = (e: HistoryEntry): void => setSettings((s) => ({ ...s, ...e.settings, dataset: s.dataset, csvText: s.csvText }));

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-100">
      <a href="#main" className="sr-only focus:not-sr-only focus:absolute focus:left-2 focus:top-2 focus:z-50 focus:rounded focus:bg-white focus:px-3 focus:py-2 focus:text-slate-900">
        Skip to content
      </a>
      <header className="border-b border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
        <div className="mx-auto flex max-w-[96rem] flex-wrap items-center justify-between gap-3 px-4 py-3">
          <div className="min-w-0">
            <h1 className="text-lg font-bold sm:text-xl">
              javascript-excel playground{" "}
              <span className="ml-1 rounded bg-amber-100 px-1.5 py-0.5 align-middle font-mono text-xs font-semibold text-amber-900 dark:bg-amber-900 dark:text-amber-100" title="The version this page runs (VERSION export)">
                v{VERSION}
              </span>
            </h1>
            <p className="text-sm text-slate-700 dark:text-slate-300">Read, write and convert .xlsx and .csv with a zero-dependency TypeScript engine. Everything runs in your browser.</p>
          </div>
          <nav aria-label="Project links" className="flex flex-wrap items-center gap-1 text-sm">
            <a className="rounded px-2 py-1 font-medium text-blue-800 hover:bg-slate-100 dark:text-blue-300 dark:hover:bg-slate-800" href={LINKS.github} target="_blank" rel="noopener noreferrer">
              GitHub
            </a>
            <a className="rounded px-2 py-1 font-medium text-blue-800 hover:bg-slate-100 dark:text-blue-300 dark:hover:bg-slate-800" href={LINKS.npm} target="_blank" rel="noopener noreferrer">
              npm
            </a>
            <a className="rounded px-2 py-1 font-medium text-blue-800 hover:bg-slate-100 dark:text-blue-300 dark:hover:bg-slate-800" href={LINKS.docs} target="_blank" rel="noopener noreferrer">
              Docs
            </a>
            <Button variant="ghost" onClick={theme.toggle} title="Toggle theme (Alt+T)" ariaKeyshortcuts="Alt+T">
              <span aria-hidden="true">{theme.dark ? "☀" : "☾"}</span>
              <span>{theme.dark ? "Light" : "Dark"}</span>
            </Button>
            <Button variant="ghost" onClick={() => setHelp(true)} title="Keyboard shortcuts (?)" ariaKeyshortcuts="?">
              <kbd aria-hidden="true" className="rounded border border-slate-300 px-1 font-mono text-xs dark:border-slate-600">?</kbd>
              <span>Shortcuts</span>
            </Button>
          </nav>
        </div>
      </header>

      <main id="main" className="mx-auto max-w-[96rem] space-y-4 px-4 py-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div role="tablist" aria-label="Playground sections" className="inline-flex rounded-lg border border-slate-300 bg-white p-1 dark:border-slate-700 dark:bg-slate-900">
            {TABS.map((t, i) => (
              <button
                key={t}
                id={`tab-${t}`}
                type="button"
                role="tab"
                aria-selected={settings.tab === t}
                aria-controls={`panel-${t}`}
                aria-keyshortcuts={`Alt+${i + 1}`}
                tabIndex={settings.tab === t ? 0 : -1}
                onClick={() => setTab(t)}
                onKeyDown={(e) => {
                  const n = TABS.length;
                  const next = e.key === "ArrowRight" ? (i + 1) % n : e.key === "ArrowLeft" ? (i - 1 + n) % n : -1;
                  if (next >= 0) {
                    e.preventDefault();
                    setTab(TABS[next]!);
                    document.getElementById(`tab-${TABS[next]}`)?.focus();
                  }
                }}
                className={`rounded-md px-4 py-1.5 text-sm font-medium ${settings.tab === t ? "bg-blue-700 text-white" : "text-slate-800 hover:bg-slate-100 dark:text-slate-200 dark:hover:bg-slate-800"}`}
              >
                {TAB_LABELS[t]}
              </button>
            ))}
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-sm">
              <input type="checkbox" className="h-4 w-4 accent-blue-700" checked={settings.worker} onChange={(e) => setSettings((s) => ({ ...s, worker: e.target.checked }))} aria-keyshortcuts="Alt+W" />
              Run in a Web Worker
            </label>
            <Button onClick={share} title="Copy a link to these settings (Alt+S)" ariaKeyshortcuts="Alt+S">
              Share link
            </Button>
            <span role="status" className="text-sm text-slate-700 dark:text-slate-300">
              {shared}
            </span>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-4 2xl:grid-cols-[minmax(0,1fr)_26rem]">
          <div id={`panel-${settings.tab}`} role="tabpanel" aria-labelledby={`tab-${settings.tab}`} className="min-w-0">
            {settings.tab === "read" && <ReadPanel settings={settings} setRead={setRead} setCsvParse={setCsvParse} onRun={onRun} fileInputRef={fileInput} runSignal={runSignal} source={source} setSource={setSource} />}
            {settings.tab === "write" && <WritePanel settings={settings} setWrite={setWrite} setDataset={setDataset} onRun={onRun} runSignal={runSignal} />}
            {settings.tab === "csv" && <CsvPanel settings={settings} setCsvText={setCsvText} setCsvParse={setCsvParse} setCsvWrite={setCsvWrite} onRun={onRun} runSignal={runSignal} />}
          </div>
          <div className="grid min-w-0 grid-cols-1 content-start gap-4 lg:grid-cols-2 2xl:grid-cols-1">
            <Card title="Examples">
              <ul className="space-y-1">
                {EXAMPLES.map((ex) => (
                  <li key={ex.id}>
                    <button type="button" onClick={() => setSettings((s) => ex.apply({ ...s }))} className="w-full rounded px-2 py-1.5 text-left hover:bg-slate-100 focus-visible:outline-2 focus-visible:outline-blue-600 dark:hover:bg-slate-800">
                      <span className="block text-sm font-medium">{ex.title}</span>
                      <span className="block text-xs text-slate-600 dark:text-slate-400">{ex.description}</span>
                    </button>
                  </li>
                ))}
              </ul>
            </Card>
            <CodePanel snippet={snippet} version={VERSION} />
            <HistoryPanel entries={runs} onRestore={restore} onClear={() => setRuns([])} />
          </div>
        </div>
      </main>

      <footer className="mx-auto max-w-[96rem] px-4 pb-8 pt-2 text-sm text-slate-700 dark:text-slate-300">
        <p>
          javascript-excel {VERSION} is an alpha on the npm <code>next</code> tag; the API can still change before 2.0.0. Files you open are read in this tab only: no uploads, no tracking, no cookies. MIT licence.{" "}
          <a className="text-blue-800 underline dark:text-blue-300" href={LINKS.playground} target="_blank" rel="noopener noreferrer">
            Playground source
          </a>
          .
        </p>
      </footer>
      <ShortcutsDialog open={help} onClose={() => setHelp(false)} shortcuts={SHORTCUTS} />
    </div>
  );
}
