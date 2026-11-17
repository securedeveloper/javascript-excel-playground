// Recent runs with their timings, and a per-action performance summary.
import { formatMs } from "../lib/engine";
import { stats, type HistoryEntry } from "../lib/history";
import { Button, Card } from "./ui";

interface Props {
  entries: HistoryEntry[];
  onRestore: (e: HistoryEntry) => void;
  onClear: () => void;
}

export function HistoryPanel({ entries, onRestore, onClear }: Props) {
  const summary = stats(entries);
  return (
    <Card title="History and timings" actions={entries.length > 0 && <Button variant="ghost" onClick={onClear}>Clear</Button>}>
      {summary.length > 0 && (
        <table className="mb-3 w-full text-left text-sm">
          <caption className="sr-only">Time per action over the recent runs</caption>
          <thead className="text-xs text-slate-600 dark:text-slate-400">
            <tr>
              <th scope="col">Action</th>
              <th scope="col" className="text-right">Runs</th>
              <th scope="col" className="text-right">Median</th>
              <th scope="col" className="text-right">Best</th>
            </tr>
          </thead>
          <tbody>
            {summary.map((s) => (
              <tr key={s.action}>
                <td className="font-mono text-xs">{s.action}</td>
                <td className="text-right">{s.runs}</td>
                <td className="text-right">{formatMs(s.median)}</td>
                <td className="text-right">{formatMs(s.best)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {entries.length === 0 ? (
        <p className="text-sm text-slate-600 dark:text-slate-400">Runs appear here. They are stored in this browser only.</p>
      ) : (
        <ol className="max-h-80 space-y-1 overflow-auto text-sm">
          {entries.map((e) => (
            <li key={e.id}>
              <button
                type="button"
                onClick={() => onRestore(e)}
                className="w-full rounded px-2 py-1 text-left hover:bg-slate-100 focus-visible:outline-2 focus-visible:outline-blue-600 dark:hover:bg-slate-800"
                title="Restore these settings"
              >
                <span className="flex justify-between gap-2">
                  <span className="truncate font-mono text-xs">
                    {e.action} · {e.source}
                  </span>
                  <span className={`shrink-0 text-xs ${e.ok ? "text-slate-700 dark:text-slate-300" : "text-red-700 dark:text-red-300"}`}>{e.ok ? formatMs(e.ms) : "failed"}</span>
                </span>
                <span className="block truncate text-xs text-slate-600 dark:text-slate-400">
                  {new Date(e.at).toLocaleTimeString()} · {e.where === "worker" ? "worker" : "main"} · {e.summary}
                </span>
              </button>
            </li>
          ))}
        </ol>
      )}
    </Card>
  );
}
