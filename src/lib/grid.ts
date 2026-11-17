// Pure helpers for the sheet preview: selection maths and render-model styles as CSS.
import { encodeCell, type CellInput, type Workbook } from "javascript-excel";
import { formatCell } from "javascript-excel/format";
import { createRenderSheet, type CellRef, type RenderFormatter, type RangeRef, type RenderBorder, type RenderCell, type RenderFont, type RenderSheet } from "javascript-excel/render";
import { useMemo, type CSSProperties } from "react";

/** The preview draws at most this window; the inspector still sees every cell. */
export const MAX_PREVIEW_ROWS = 200;
export const MAX_PREVIEW_COLS = 30;
export const HEADER_H = 22;
export const ROW_HEADER_W = 44;

export interface Selection {
  anchor: CellRef;
  focus: CellRef;
}

export function selectionRange(sel: Selection): RangeRef {
  return {
    start: { row: Math.min(sel.anchor.row, sel.focus.row), col: Math.min(sel.anchor.col, sel.focus.col) },
    end: { row: Math.max(sel.anchor.row, sel.focus.row), col: Math.max(sel.anchor.col, sel.focus.col) },
  };
}

export function rangeLabel(r: RangeRef): string {
  const a = encodeCell(r.start);
  const b = encodeCell(r.end);
  return a === b ? a : `${a}:${b}`;
}

/**
 * formatCell, plus a date format for Date values that have none (CSV and
 * built datasets without numFmt), so they show as dates instead of serials.
 * Display only: the model is not changed.
 */
export const displayFormat: RenderFormatter = (cell, options) =>
  cell.value instanceof Date && cell.style?.numFmt === undefined ? formatCell({ ...cell, style: { ...cell.style, numFmt: "yyyy-mm-dd" } }, options) : formatCell(cell, options);

/** Creates the render handle for a sheet, with Excel-style formatted text. */
export function useRenderSheet(workbook: Workbook<CellInput>, sheet: number, locale: string): RenderSheet | undefined {
  return useMemo(() => {
    if (!workbook.sheets[sheet]) return undefined;
    return createRenderSheet(workbook, sheet, { format: displayFormat, locale, margin: { rows: 3, cols: 2 } });
  }, [workbook, sheet, locale]);
}

export function cssBorder(b: RenderBorder | undefined): string | undefined {
  if (!b || b.style === "none") return undefined;
  const map: Record<string, string> = {
    thin: "1px solid",
    hair: "1px dotted",
    dotted: "1px dotted",
    dashed: "1px dashed",
    dashDot: "1px dashed",
    dashDotDot: "1px dashed",
    medium: "2px solid",
    mediumDashed: "2px dashed",
    mediumDashDot: "2px dashed",
    mediumDashDotDot: "2px dashed",
    slantDashDot: "2px dashed",
    thick: "3px solid",
    double: "3px double",
  };
  return `${map[b.style] ?? "1px solid"} ${b.color}`;
}

export function fontCss(f: RenderFont): CSSProperties {
  return {
    fontFamily: `"${f.family}", Calibri, Carlito, "Segoe UI", Arial, sans-serif`,
    fontSize: `${f.sizePt}pt`,
    color: f.color,
    fontWeight: f.bold ? 700 : 400,
    fontStyle: f.italic ? "italic" : "normal",
    textDecoration: [f.underline !== "none" ? "underline" : "", f.strike ? "line-through" : ""].filter(Boolean).join(" ") || "none",
    verticalAlign: f.vertAlign === "superscript" ? "super" : f.vertAlign === "subscript" ? "sub" : undefined,
  };
}

export function cellCss(cell: RenderCell): CSSProperties {
  const s = cell.style;
  const css: CSSProperties = { ...fontCss(s.font) };
  delete css.verticalAlign;
  if (cell.color) css.color = cell.color;
  const bg = s.background;
  if (bg.kind === "solid") css.backgroundColor = bg.color;
  else if (bg.kind === "pattern") css.backgroundColor = bg.pattern === "none" ? undefined : bg.fg;
  else if (bg.kind === "gradient" && bg.stops.length) {
    css.backgroundImage = `linear-gradient(${bg.degree + 90}deg, ${bg.stops.map((st) => `${st.color} ${Math.round(st.position * 100)}%`).join(", ")})`;
  }
  const h = s.align.horizontal;
  css.textAlign = h === "general" || h === "fill" ? cell.align : h === "centerContinuous" ? "center" : h === "distributed" ? "justify" : h;
  css.verticalAlign = s.align.vertical === "center" ? "middle" : s.align.vertical === "top" ? "top" : "bottom";
  if (s.align.wrap) css.whiteSpace = "pre-wrap";
  if (s.align.indent) css.paddingInlineStart = `${3 + s.align.indent * 9}px`;
  const e = cell.edges;
  const t = cssBorder(e.top);
  const l = cssBorder(e.left);
  const r = cssBorder(e.right);
  const b = cssBorder(e.bottom);
  if (t) css.borderTop = t;
  if (l) css.borderLeft = l;
  if (r) css.borderRight = r;
  if (b) css.borderBottom = b;
  return css;
}

