// Recent runs, kept in localStorage. Only settings and summaries are stored,
// never file contents.
import type { Settings, Tab } from "./state";

export interface HistoryEntry {
  id: string;
  at: number;
  tab: Tab;
  /** What ran: a file name, "sample", "dataset" or "CSV text". */
  source: string;
  action: string;
  ms: number;
  where: "main" | "worker";
  summary: string;
  ok: boolean;
  settings: Omit<Settings, "dataset" | "csvText">;
}

const KEY = "javascript-excel-playground:history:v1";
export const MAX_HISTORY = 25;
/** Runs of the same action closer together than this replace each other. */
export const REPEAT_MS = 1500;

function storage(): Storage | undefined {
  try {
    return typeof localStorage === "undefined" ? undefined : localStorage;
  } catch {
    return undefined;
  }
}

export function loadHistory(): HistoryEntry[] {
  try {
    const raw = storage()?.getItem(KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? (parsed as HistoryEntry[]).filter((e) => e && typeof e.id === "string" && typeof e.ms === "number").slice(0, MAX_HISTORY) : [];
  } catch {
    return [];
  }
}

export function saveHistory(entries: HistoryEntry[]): void {
  try {
    storage()?.setItem(KEY, JSON.stringify(entries.slice(0, MAX_HISTORY)));
  } catch {
    // Storage full or blocked: history is a convenience, so ignore.
  }
}

export function addEntry(entries: HistoryEntry[], entry: Omit<HistoryEntry, "id" | "at">, now = Date.now()): HistoryEntry[] {
  const full: HistoryEntry = { ...entry, id: `${now.toString(36)}-${Math.random().toString(36).slice(2, 8)}`, at: now };
  // Typing in an editor re-runs on every change: fold quick repeats of the same action into one entry.
  const last = entries[0];
  const rest = last && last.action === entry.action && last.source === entry.source && last.tab === entry.tab && now - last.at < REPEAT_MS ? entries.slice(1) : entries;
  return [full, ...rest].slice(0, MAX_HISTORY);
}

/** Median and fastest time per action, for the performance summary. */
export function stats(entries: HistoryEntry[]): Array<{ action: string; runs: number; median: number; best: number }> {
  const by = new Map<string, number[]>();
  for (const e of entries) if (e.ok) by.set(e.action, [...(by.get(e.action) ?? []), e.ms]);
  return [...by.entries()].map(([action, times]) => {
    const sorted = [...times].sort((a, b) => a - b);
    const mid = sorted.length >> 1;
    const median = sorted.length % 2 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2;
    return { action, runs: sorted.length, median, best: sorted[0]! };
  });
}
