# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.2.1] - 2026-11-11

### Changed

- Faster XLSX reads: short cell tags are scanned with a byte loop instead of
  `Buffer.indexOf`, and decimal values use a precomputed powers-of-ten table.
  Reads are about 25% faster on the mixed and number-heavy benchmark files.

## [3.2.0] - 2026-11-08

### Added

- Shared writer core (`src/writers/core`) used by the Node and browser XLSX/XLSB
  writers. The unpacked package parts are byte-identical between Node and browsers.
- `useSharedStrings: false` writer option for inline strings (four writers).
- `ExcelError` cell value and round-trip support for Excel error cells.
- `XlsxReader.openBuffer()` and `XlsbReader.openBuffer()` plus `readRow()`,
  `readAll()` and `rows()` batch APIs.
- Isolated-process I/O benchmark (`npm run benchmark:io`) with a 3% regression
  gate, and a self-contained Microsoft Excel COM conformance suite
  (`EXCEL_COM=1 npm run test:excel-com`).
- Etap 2 cell formatting: `CellStyle`, `StyledCell`, `FormulaCell`, `RichText`
  and style helpers. `FormattedCell` remains supported.
- XLSX styling: fonts, fills, borders, alignment, protection, custom number
  formats and differential styles for conditional formatting.
- XLSX sheet features: formula cells with cached values, rich text, merged
  cells, hyperlinks, data validation, conditional formatting and cell comments
  (classic notes with VML).
- XLSB styling and formulas: a BIFF12 style registry and a curated formula
  compiler with cached values; merged cells, hyperlinks and data validation.
- Conservative column autofit (header + up to 200 sample rows, capped at 200,
  not narrower than xlspy/Excel AutoFit up to that cap) and explicit
  `columnWidths` overrides in all four writers. Streaming writers delay the
  sheet head until the first 200 data rows have been measured.
- `XlsxReader.getStyle()`, `getFormula()` and `getRichText()`; `XlsbReader.getStyle()`.

### Changed

- `writeSheet` accepts a sheet-options object in place of the legacy
  `doAutofilter` boolean while keeping the boolean overload backward compatible.
- `updateColWidthsFromRows` understands styled, formula and rich-text cells.
- Autofit is deliberately conservative so it is not narrower than
  xlspy/Excel AutoFit up to the 200 cap: columns use the 8.43 minimum, calibrated
  character costs (CJK/fullwidth, emoji, wide fonts, bold/italic and font
  size), number/date-format display text, room for the autofilter button, the
  5/7-character cell padding, and a 200-character cap on the stored width. The
  default fallback width is 9.14.
- Autofit also measures alignment indent, the longest line of wrapped text,
  negative number-format sections, and Excel's 11-character General rendering
  instead of the full JavaScript decimal expansion.
- `options.autofitRows` controls how many rows are measured (default 200;
  `0` disables row sampling).
- XLSB column widths are written as full 1/256-character units instead of
  floored integers.
- Explicit `columnWidths` values are validated to 0.5-255 characters.

### Fixed

- XLSX reading no longer scans the rest of the sheet for every cell. The
  formula lookup was unbounded, so a 20,000 x 10 sheet took minutes to read
  instead of milliseconds; tag and value lookups are now bounded to the cell.
- XLSX numeric cells are parsed without allocating an intermediate string
  (`parseFloat` is still used for exponents and long mantissas), restoring
  read throughput close to the pre-Etap-2 reader.
- Dates with a time of day now use the datetime style (`numFmt` 22) in XLSX and
  XLSB instead of the date-only style, so the time is no longer hidden.
- Streaming `writeRow()` now feeds autofit instead of leaving columns sized only
  from the header and `sampleRows`.
- Long runs of accented Latin, Greek/Cyrillic and symbol characters are no
  longer measured narrower than xlspy; a per-character exception table covers
  the outliers.
- XLSB `BrtWsDim` now contains the real last row and an inclusive last column;
  previously rows were always 0 and the column was one past the data.
- XLSB cells with an error value are read as `#DIV/0!`-style text instead of `null`.
- XLSX reader resolves optional `r`/`t`/`s` cell attributes from raw bytes
  (previously attribute-less and typed cells could be misread).
- XLSB custom number formats no longer produce a corrupt styles part.

### Limitations

- XLSB rich text, comments and conditional formatting are not emitted yet; the
  writers throw a clear error instead of silently dropping them.
- XLSB formula support is a subset (literals, A1 references/ranges, operators
  and common worksheet functions); unsupported syntax throws.
- `XlsbReader` does not reconstruct formula text; it returns cached values.

## [3.1.1] - 2026-10-24

### Changed

- The default ZIP compression level is 3 (was 6) for XLSX, XLSB and the
  streaming updater. Level 6 made XLSB writes about 2.6x slower for an 8%
  smaller file.

## [3.1.0] - 2026-10-24

### Changed

- XLSX output omits the `r` attribute on cells that directly follow the previous
  cell (`<c><v>1</v></c>`). A cell keeps `r` after a gap, at the start of a row that
  does not begin in column A, and in the header row only when needed. Excel and the
  readers in this package resolve the position from the order of cells.
- XLSX and XLSB writers accept `{ compressionLevel }` (integer 0–9, default 6). The
  archiver default was level 1, which compresses less.

### Fixed

- Raw CRLF and lone CR in XML text (shared strings, inline strings, `str` cells) are
  normalized to LF on read, as XML 1.0 requires. `&#13;` still reads as CR.
- `escapeXmlText` writes CR as `&#13;`, so a CR written by the library survives a
  round trip instead of being normalized away. The XLSX writer now does the same in
  shared strings, where a raw CR was previously written and read back as LF by Excel.
- Text containing `_xHHHH_` is escaped as `_x005F_xHHHH_` when written, and
  `_xHHHH_` written by Excel (for example `_x000D_`) is decoded on read (ECMA-376
  ST_Xstring). Attribute values are not affected.
- `XlsxUpdater` infers the column of a cell without `r` when it reads existing styles,
  so style inheritance works for files produced by this library and by xlspy.

## [3.0.0] - 2026-10-23

### Breaking

- `XlsxReader` / `XlsbReader` now return an empty row (`[]`, `getValue()` → `null`)
  for every missing row position, including leading gaps, based on the row `r`
  attribute (XLSX) and `BrtRowHdr` (XLSB). Like xlspy, gap positions are yielded;
  unlike xlspy's padded `[None] * cols`, this reader yields a zero-length row.
  Code that counts rows must account for gap rows.

### Documentation

- Date conversion is UTC-based on both write and read: a `Date` is stored by its
  UTC instant (`getTime()`), so `new Date(2024, 5, 30)` in UTC+2 appears as
  2024-06-29 22:00 in Excel. Use `Date.UTC(...)` for date-only values.

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
