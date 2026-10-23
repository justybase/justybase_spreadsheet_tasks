# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Breaking

- `XlsxReader` / `XlsbReader` now return an empty row (`[]`, `getValue()` → `null`)
  for every missing row position, including leading gaps, based on the row `r`
  attribute (XLSX) and `BrtRowHdr` (XLSB). Like xlspy, gap positions are yielded;
  unlike xlspy's padded `[None] * cols`, this reader yields a zero-length row.
  Code that counts rows must account for gap rows.

### Added

- `selectSheet(nameOrIndex)` and `nextResult()` on `ExcelReaderAbstract`,
  implemented by `XlsxReader` and `XlsbReader`. Default remains sheet 0; state
  (row/EOF, `fieldCount`) resets on selection.
- Excel limits (`MAX_ROWS=1_048_576`, `MAX_COLS=16_384`, `MAX_CELL_CHARS=32_767`
  UTF-16 units) enforced in `XlsxWriter`, `XlsbWriter`, `BrowserXlsxWriter` and
  `BrowserXlsbWriter`. Overlong text, oversized dimensions and rows/headers
  longer than `columnCount` throw instead of writing corrupt files (xlspy parity).
- Shared XML helpers: `parseTagAttributes()`, `resolveRelTarget()` and
  single-pass `unescapeXml()` in `src/xmlUtils.ts`, exported from the package
  entry point.
- `convertXlsbToXlsx` and `convertXlsxToXlsb` — Windows/Excel COM conversions
  with protected destinations and optional explicit overwriting.
- `replaceSheetDataStream()` and `saveStreaming()` for XLSX/XLSB updaters,
  accepting synchronous or asynchronous row sources without collecting the
  complete result set in memory.
- `XlsmUpdater` as an explicit alias for `XlsxUpdater`; VBA package members are
  preserved as opaque ZIP parts.

### Fixed

- `BrowserXlsxWriter` / `BrowserXlsbWriter` skip a `FormattedCell` whose value is
  `null` or `undefined`, matching the Node writers (previously a `TypeError`).
- Browser sources in `src/browser/` are type-checked under `strict` by
  `tsconfig.browser.json`, which `npm run typecheck` runs. The `@ts-nocheck` pragmas
  were removed.
- Column letters ≥702 (`AAA`…`XFD`): `XlsxWriter` and `BrowserXlsxWriter` now
  delegate to `columnIndexToLetter()` instead of falling back to `'A'`.
- XLSB shared-string table: `BrtBeginSst` now writes `total` then `unique`
  (was swapped); `BiffReaderWriter` and updaters already expected that order.
- Writer finalization is atomic: output goes to a same-directory temp file,
  output/archive errors reject the promise, streams are destroyed and the temp
  file removed; only a complete workbook is renamed into place (xlspy parity).
  `writeAdmZipStreaming()` now observes output errors before/during the append loop.
- Workbook relationships and sheet paths: absolute `Target="/xl/…"` resolved
  correctly, attribute order no longer matters (`parseTagAttributes` used in
  both readers and both updaters). Verified against openpyxl output.
- `unescapeXml()` single-pass: `&amp;lt;` stays `&lt;` instead of becoming `<`.
- Shared strings: `<rPh>` phonetics skipped, `<si/>` yields `""`, rich-text
  `inlineStr` concatenates all `<t>` runs.
- XLSX cell types: `t="str"` → text, `t="e"` → error text, `t="d"` → `Date`.
- Self-closing `<row …/>` no longer swallows the next row.
- 1904 date system: `workbookPr date1904` (XLSX) and `BrtWbProp` (XLSB) set the
  reader/updater epoch (`1904-01-01` vs `1899-12-30`).
- XLSB dimension: updaters patch `BrtWsDim` (`0x94`, `rwLast@4`/`colLast@12`)
  instead of `BrtSel` (`0x98`).
- Stale parts after `replaceSheetData`: `xl/calcChain.xml` (`.bin` for XLSB)
  removed with its `Content_Types`/rels entries; XLSX `xl/tables/*.xml` refs
  (`<table>` + `<autoFilter>`) updated to the new dimension when the table's
  width is unchanged, without expanding over other tables. XLSB tables are not
  patched (documented limitation).
- Destination replacement during Excel conversion now preserves files created
  concurrently and rolls back safely when the final rename fails.
- Shared-string `count` values now reflect the replaced worksheet rather than
  accumulating stale references across repeated updates. Pending strings are
  committed only once.
- Updater saves now install the destination atomically through a same-directory
  temporary file.
- 1900 date system leap-day gap: Excel counts a fictitious 1900-02-29, so
  serials 1-59 denote 1900-01-01..1900-02-28. `oaDateToTime()` /
  `timeToOaDate()` in `writerHelpers.ts` apply the one-day shift (both
  directions, all readers/writers/updaters incl. browser builds); the 1904
  system is unaffected. Previously serial 1 read as 1899-12-31, disagreeing
  with Excel, xlspy and openpyxl.
- `XlsbUpdater` date fallback now recognises custom `BrtFmt` date formats
  (e.g. xlspy's `yyyy-mm-dd`) via `BiffReaderWriter.readStyles()` instead of
  only built-in numFmt ranges, so replacement `Date` cells keep a date style
  instead of degrading to raw serials.

## [2.1.0] - 2026-08-07

### Added

- `XlsxUpdater` and `XlsbUpdater` — replace the data of a worksheet inside an
  existing `.xlsx` / `.xlsb` file without rebuilding the workbook. Pivot tables,
  other sheets, styles and defined names are preserved; shared strings are
  appended (never reordered), column styles are inherited, and pivot cache
  ranges + `refreshOnLoad` are updated automatically.
- Shared helpers: `xmlUtils.ts` (XML escaping, column letters, shared strings
  parsing) and `biff12Utils.ts` (BIFF12 record helpers for XLSB).
- `examples/update-existing.ts` — end-to-end example for refreshing a data
  sheet in an existing XLSX / XLSB report.

### Fixed

- `XlsbUpdater` pivot source range was one row too long when `headers` were not
  provided, causing an empty trailing row (and a "(blank)" item in pivot
  tables). The range now ends exactly at the last data row.
- Both updaters now trim trailing rows that contain only empty cells (padding
  rows from SQL exports), so they don't appear as "(blank)" in pivot tables.
- `XlsbUpdater` no longer removes the worksheet's row-block terminator records
  (0x92/0x217/0x1dd/0x1dc) and keeps the sheet dimension record (0x98) in sync
  with the new row count — without these, Excel hid the last written row.

## [2.0.0] - 2026-07-25

### Breaking

- `XlsbReader.read()` is now `async` and returns `Promise<boolean>` (same as `XlsxReader`).
  Update loops from `while (reader.read())` to `while (await reader.read())`.
- Package entry points moved under `dist/cjs` and `dist/esm` with an `exports` map.
  Most consumers of `@justybase/spreadsheet-tasks` need no import path changes.

### Added

- `CellValue` / `PrimitiveCellValue` types for writer and reader cell APIs
- Shared writer helpers: `sanitizeSheetName`, column-width helpers, `StreamingSheetState`
- Dual package build (CJS + ESM) via `tsup`
- Browser sources under `src/browser/` with `npm run build:browser`
- `docs/PACKAGING.md` (ZIP library rationale + dual package notes)
- Dist surface smoke test (`npm run test:dist`)

### Changed

- `XlsbWriter` / `XlsxWriter` share sanitization, autofit, and streaming state helpers
- Tests import from `src/` (published surface still verified via `test:dist`)
- ZIP type definitions use `@types/*` only (hand-written ambient modules removed)

### Browser

- Source of truth is TypeScript in `src/browser/`; `browser/*.js` writer sources are generated bundles

## [1.3.2] - previous

See git history for 1.x changes.
