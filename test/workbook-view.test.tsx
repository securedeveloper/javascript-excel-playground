// @vitest-environment happy-dom
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { read } from "javascript-excel/read";
import { afterEach, describe, expect, it } from "vitest";
import { WorkbookView } from "../src/components/WorkbookView";
import { sampleXlsx } from "../src/lib/samples";

afterEach(cleanup);

describe("WorkbookView", () => {
  it("draws the sample, switches sheets and inspects a cell", async () => {
    const { workbook, warnings } = await read(await sampleXlsx(), { dates: "serial" });
    render(<WorkbookView workbook={workbook} warnings={warnings} locale="en-US" />);
    expect(screen.getAllByText("Quarterly orders").length).toBeGreaterThan(0);
    expect(screen.getByText("$777.00")).toBeTruthy();
    fireEvent.mouseDown(screen.getByText("$777.00"));
    expect(screen.getByText("=D3*E3")).toBeTruthy();
    fireEvent.click(screen.getByRole("tab", { name: /Notes/ }));
    expect(screen.getByText("javascript-excel on npm").closest("a")?.getAttribute("rel")).toBe("noopener noreferrer");
  });

  it("renders cell text as text, never as HTML", () => {
    const payload = '<img src=x onerror="alert(1)">';
    const { container } = render(<WorkbookView workbook={{ sheets: [{ name: "X", rows: [[payload]] }] }} warnings={[]} locale="en-US" />);
    expect(container.querySelector("img")).toBeNull();
    expect(screen.getAllByText(payload).length).toBeGreaterThan(0);
  });
});
