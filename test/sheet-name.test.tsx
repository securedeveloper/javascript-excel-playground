// @vitest-environment happy-dom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { WritePanel } from "../src/components/WritePanel";
import { buildSheet } from "../src/lib/engine";
import { CSV_SAMPLES, SAMPLE_DATASET_JSON } from "../src/lib/samples";
import { codePointLength, sheetNameError } from "../src/lib/sheet-name";
import { decodeSettings, defaultSettings, encodeSettings, type Settings } from "../src/lib/state";

afterEach(cleanup);

const D = defaultSettings(SAMPLE_DATASET_JSON, CSV_SAMPLES[0]!.text);
// 31 code points but 32 UTF-16 units: valid for Excel and the library, too long for maxLength=31.
const EMOJI_31 = `${"x".repeat(30)}😀`;
const EMOJI_32 = `${"x".repeat(31)}😀`;

describe("sheet names are counted in code points", () => {
  it("accepts 30 × x + 😀 (31 code points, 32 UTF-16 units)", () => {
    expect(EMOJI_31.length).toBe(32);
    expect(codePointLength(EMOJI_31)).toBe(31);
    expect(sheetNameError(EMOJI_31)).toBeNull();
    expect(buildSheet(SAMPLE_DATASET_JSON, { ...D.write, sheetName: EMOJI_31 }).result.sheet.name).toBe(EMOJI_31);
  });

  it("rejects 32 code points with the library's reason", () => {
    expect(sheetNameError(EMOJI_32)).toMatch(/at most 31 characters; this one has 32/);
    expect(() => buildSheet(SAMPLE_DATASET_JSON, { ...D.write, sheetName: EMOJI_32 })).toThrow();
    expect(sheetNameError("a:b")).toMatch(/can't contain/);
    expect(sheetNameError("")).toMatch(/empty/);
  });

  it("keeps a 31-code-point name in the shared URL", () => {
    const s: Settings = { ...D, tab: "write", write: { ...D.write, sheetName: EMOJI_31 } };
    expect(decodeSettings(encodeSettings(s, D), D).write.sheetName).toBe(EMOJI_31);
    const tooLong: Settings = { ...D, write: { ...D.write, sheetName: EMOJI_32 } };
    expect(decodeSettings(encodeSettings(tooLong, D), D).write.sheetName).toBe(D.write.sheetName);
  });

  it("lets the field take the name and shows the library's verdict", () => {
    let settings: Settings = { ...D, tab: "write" };
    const noop = () => {};
    const view = (): ReturnType<typeof render> =>
      render(<WritePanel settings={settings} setWrite={(w) => (settings = { ...settings, write: { ...settings.write, ...w } })} setDataset={noop} onRun={noop} runSignal={0} />);
    let r = view();
    const input = screen.getByLabelText("Sheet name") as HTMLInputElement;
    expect(input.hasAttribute("maxlength")).toBe(false);
    fireEvent.change(input, { target: { value: EMOJI_31 } });
    expect(settings.write.sheetName).toBe(EMOJI_31);
    r.unmount();
    r = view();
    expect(screen.getByText("31 of 31 characters")).toBeTruthy();
    expect(screen.getByLabelText("Sheet name").getAttribute("aria-invalid")).toBeNull();
    r.unmount();
    settings = { ...settings, write: { ...settings.write, sheetName: EMOJI_32 } };
    view();
    expect(screen.getByLabelText("Sheet name").getAttribute("aria-invalid")).toBe("true");
    expect(screen.getByText(/at most 31 characters; this one has 32/)).toBeTruthy();
  });
});
