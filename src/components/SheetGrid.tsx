// A spreadsheet preview drawn from javascript-excel/render: geometry, merges,
// concrete styles, resolved border edges, formatted text and navigation.
// Cell text is rendered by React as text, never as HTML: files are untrusted.
import { encodeCol } from "javascript-excel";
import { colLeft, colWidth, edge, isColHidden, isRowHidden, mergeAt, renderCell, rowHeight, rowTop, sheetExtent, sheetSettings, step, type CellRef, type RenderCell, type RenderDirection, type RenderSheet } from "javascript-excel/render";
import type { CSSProperties, KeyboardEvent } from "react";
import { HEADER_H, MAX_PREVIEW_COLS, MAX_PREVIEW_ROWS, ROW_HEADER_W, cellCss, fontCss, selectionRange, type Selection } from "../lib/grid";

interface GridProps {
  view: RenderSheet;
  selection: Selection;
  onSelect: (sel: Selection) => void;
  label: string;
}

export function SheetGrid({ view, selection, onSelect, label }: GridProps) {
  const extent = sheetExtent(view);
  const settings = sheetSettings(view);
  const rows = Math.min(extent.rows, MAX_PREVIEW_ROWS);
  const cols = Math.min(extent.cols, MAX_PREVIEW_COLS);
  const range = selectionRange(selection);
  const inRange = (r: number, c: number): boolean => r >= range.start.row && r <= range.end.row && c >= range.start.col && c <= range.end.col;

  const move = (to: CellRef, extend: boolean): void => {
    const clamped = { row: Math.min(to.row, rows - 1), col: Math.min(to.col, cols - 1) };
    onSelect(extend ? { anchor: selection.anchor, focus: clamped } : { anchor: clamped, focus: clamped });
  };

  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>): void => {
    const dirs: Record<string, RenderDirection> = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right" };
    const dir = dirs[e.key];
    if (!dir) return;
    e.preventDefault();
    const from = selection.focus;
    move((e.ctrlKey || e.metaKey ? edge : step)(view, from, dir), e.shiftKey);
  };

  const colIdx = Array.from({ length: cols }, (_, c) => c).filter((c) => !isColHidden(view, c));
  const rowIdx = Array.from({ length: rows }, (_, r) => r).filter((r) => !isRowHidden(view, r));
  const freeze = settings.freeze;

  const stickyCell = (r: number, c: number): CSSProperties => {
    const css: CSSProperties = {};
    if (r < freeze.rows) {
      css.position = "sticky";
      css.top = HEADER_H + rowTop(view, r);
      css.zIndex = 2;
    }
    if (c < freeze.cols) {
      css.position = "sticky";
      css.left = ROW_HEADER_W + colLeft(view, c);
      css.zIndex = r < freeze.rows ? 3 : 2;
    }
    return css;
  };

  return (
    <div
      className="sheet-scroll relative max-h-[60vh] overflow-auto rounded-md border border-slate-300 bg-white outline-none focus-visible:ring-2 focus-visible:ring-blue-600 dark:border-slate-600"
      tabIndex={0}
      role="region"
      aria-label={`${label}. Arrow keys move, Shift extends, Ctrl or Cmd jumps to the data edge.`}
      onKeyDown={onKeyDown}
      dir={settings.rightToLeft ? "rtl" : "ltr"}
    >
      <table role="grid" aria-label={label} aria-rowcount={extent.rows + 1} aria-colcount={extent.cols + 1} aria-multiselectable="true" className={`sheet ${settings.showGridLines ? "grid-lines" : ""}`} style={{ width: ROW_HEADER_W + colIdx.reduce((s, c) => s + colWidth(view, c), 0) }}>
        <colgroup>
          <col style={{ width: ROW_HEADER_W }} />
          {colIdx.map((c) => (
            <col key={c} style={{ width: colWidth(view, c) }} />
          ))}
        </colgroup>
        <thead>
          <tr style={{ height: HEADER_H }}>
            <th className="corner" scope="col">
              <span className="sr-only">Row and column headers</span>
            </th>
            {colIdx.map((c) => (
              <th key={c} scope="col" className={c >= range.start.col && c <= range.end.col ? "hdr-active" : undefined}>
                {encodeCol(c)}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rowIdx.map((r) => (
            <tr key={r} style={{ height: rowHeight(view, r) }}>
              <th scope="row" className={r >= range.start.row && r <= range.end.row ? "hdr-active" : undefined} style={r < freeze.rows ? { position: "sticky", top: HEADER_H + rowTop(view, r), zIndex: 3 } : undefined}>
                {r + 1}
              </th>
              {colIdx.map((c) => {
                const m = mergeAt(view, r, c);
                if (m && (m.anchor.row !== r || m.anchor.col !== c)) return null;
                const cell = renderCell(view, r, c);
                const selected = inRange(r, c);
                const focus = selection.focus.row === r && selection.focus.col === c;
                const span = m ? { rowSpan: Math.min(m.range.end.row, rows - 1) - r + 1, colSpan: Math.min(m.range.end.col, cols - 1) - c + 1 } : {};
                return (
                  <td
                    key={c}
                    {...span}
                    aria-selected={selected}
                    className={`${selected ? "sel" : ""} ${focus ? "focus" : ""} ${cell?.style.align.wrap ? "" : "nowrap"}`}
                    style={{ ...(cell ? cellCss(cell) : {}), ...stickyCell(r, c) }}
                    title={cell?.formula ?? cell?.hyperlink?.tooltip ?? undefined}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      (e.currentTarget.closest("[role=region]") as HTMLElement | null)?.focus();
                      move({ row: r, col: c }, e.shiftKey);
                    }}
                  >
                    {cell ? <CellContent cell={cell} /> : null}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CellContent({ cell }: { cell: RenderCell }) {
  if (cell.runs?.length) {
    return (
      <>
        {cell.runs.map((run, i) => (
          <span key={i} style={fontCss(run.font)}>
            {run.text}
          </span>
        ))}
      </>
    );
  }
  const text = cell.fill && !cell.text ? cell.fill.char.repeat(8) : cell.text;
  const link = cell.hyperlink;
  if (link?.kind === "external" && link.allowed) {
    return (
      <a href={link.href} target="_blank" rel="noopener noreferrer" style={{ color: "inherit" }} onMouseDown={(e) => e.stopPropagation()}>
        {text}
      </a>
    );
  }
  return <>{text}</>;
}
