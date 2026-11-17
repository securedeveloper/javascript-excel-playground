import { describe, expect, it } from "vitest";
import { addEntry, REPEAT_MS, stats, type HistoryEntry } from "../src/lib/history";
import { decodeSettings, defaultSettings, encodeSettings, MAX_URL_TEXT } from "../src/lib/state";

const D = defaultSettings('{"columns":[],"data":[]}', "a,b");

describe("URL state", () => {
  it("is empty for the defaults and round-trips changes", () => {
    expect(encodeSettings(D, D)).toBe("");
    const s = { ...D, tab: "csv" as const, worker: true, read: { ...D.read, dates: "serial" as const, maxRows: 7 }, csvParse: { ...D.csvParse, delimiter: "\t" as const }, csvText: "x;y\nÄ;€" };
    const q = encodeSettings(s, D);
    expect(q).toMatch(/^tab=csv&s=[\w-]+$/);
    expect(decodeSettings(q, D)).toEqual(s);
  });

  it("leaves long texts out of the URL", () => {
    const s = { ...D, dataset: "x".repeat(MAX_URL_TEXT + 1) };
    expect(decodeSettings(encodeSettings(s, D), D).dataset).toBe(D.dataset);
  });

  it("treats the URL as untrusted input", () => {
    const evil = btoa(JSON.stringify({ worker: "yes", read: { maxRows: -5, dates: "<script>", locale: "xx" }, write: { headerFill: "red; background:url(x)" }, __proto__: { polluted: true } }));
    const s = decodeSettings(`tab=nope&s=${evil}`, D);
    expect(s).toEqual(D);
    expect(decodeSettings("s=%%%not-base64", D)).toEqual(D);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

describe("history", () => {
  const entry = (action: string): Omit<HistoryEntry, "id" | "at"> => ({ tab: "read", source: "sample", action, ms: 5, where: "main", summary: "", ok: true, settings: { ...D } });

  it("folds quick repeats and keeps distinct runs", () => {
    let h = addEntry([], entry("read"), 1000);
    h = addEntry(h, entry("read"), 1000 + REPEAT_MS - 1);
    expect(h).toHaveLength(1);
    h = addEntry(h, entry("read"), 1000 + 3 * REPEAT_MS);
    h = addEntry(h, entry("parseCsv"), 1000 + 3 * REPEAT_MS);
    expect(h.map((e) => e.action)).toEqual(["parseCsv", "read", "read"]);
  });

  it("summarises timings per action", () => {
    const h = [3, 1, 2].map((ms, i) => ({ ...entry("read"), ms, id: String(i), at: i }));
    expect(stats(h)).toEqual([{ action: "read", runs: 3, median: 2, best: 1 }]);
  });
});
