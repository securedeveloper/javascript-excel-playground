# javascript-excel playground

An interactive playground for [javascript-excel](https://www.npmjs.com/package/javascript-excel), the zero-dependency TypeScript engine that reads and writes `.xlsx` and `.csv`. It is a single-page app with no backend: files you open are read in your browser tab and never leave your machine.

It installs the **published** package from npm at an exact version (currently `2.0.0`, the `latest` release), so what you see is what `npm install` gives you.

## What it shows

- **Read**: open an `.xlsx` or `.csv` (drag and drop, or the built-in sample) and see sheets, cells, styles, merges, freeze panes, autofilter, links and the reader's warnings. Options: styles, formulas, dates as `Date` or serials, rich text, `maxRows`, and the CSV parse options.
- **Render model preview**: the grid is drawn from `javascript-excel/render` with `formatCell` from `javascript-excel/format`: displayed text in a chosen locale, concrete colours, resolved border edges, merges, frozen rows, Ctrl/Cmd+arrow navigation, a cell inspector (model cell, formula, link, resolved style), selection stats and TSV/HTML copy.
- **Write**: edit a `{ columns, data }` dataset as JSON, or an array of them to place several blocks on one sheet with `fromDatasets` (below each other, or side by side with `placement: 'right'`), tune the options (freeze, autofilter, auto width, totals, header fill, banded rows, gap) and compression, preview the sheet, download `.xlsx` or `.csv`, or write and read it back.
- **CSV**: parse text with delimiter detection, type inference, strict or lenient mode, see typed rows and warnings, and serialise it back with `toCsv` options, including the formula-injection guard.
- **Web Worker mode**: one switch runs reading, writing and CSV work through `javascript-excel/worker`.
- **Shared tools**: one-click examples, a code generator that writes the snippet for the current settings (with install commands and copy buttons), shareable URLs, a local history of runs with timings, light and dark themes and keyboard shortcuts (press `?`).

The sample workbook is generated with the library at runtime (`src/lib/samples.ts`), so the repository holds no binary fixtures.

## Local development

Requires Node 24 or newer.

```sh
npm ci
npm run dev        # http://localhost:5173
npm run typecheck
npm run lint
npm test
npm run build      # outputs dist/
npm run preview    # serves dist/ with the same security headers as production
```

## Deploy

The site is deployed on Vercel from `main`: import the repository with the **Vite** framework preset and the default build command and output directory. `vercel.json` only adds security headers (a strict Content-Security-Policy, `nosniff`, a referrer policy and `frame-ancestors 'none'`); `vite preview` serves the same headers, so they can be checked locally.

To move to a new library release, bump the exact `javascript-excel` version in `package.json`, run `npm install`, and check the playground with `npm test` and `npm run build`.

## License

MIT, see [LICENSE](LICENSE).
