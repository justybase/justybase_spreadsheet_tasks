# API Documentation

## Table of Contents

- [XlsbWriter](#xlsbwriter)
- [XlsxWriter](#xlsxwriter)
- [XlsbReader](#xlsbreader)
- [XlsxReader](#xlsxreader)
- [XlsxUpdater](#xlsxupdater)
- [XlsmUpdater](#xlsmupdater)
- [XlsbUpdater](#xlsbupdater)
- [Excel conversion](#excel-conversion)
- [ReaderFactory](#readerfactory)
- [Types](#types)

---

## XlsbWriter

High-performance writer for Excel Binary Workbook (.xlsb) format.

### Constructor

```typescript
new XlsbWriter(filePath: string)
```

**Parameters:**
- `filePath` - Path where the XLSB file will be created

**Example:**
```typescript
const writer = new XlsbWriter('output.xlsb');
```

### Methods

#### `addSheet(sheetName: string, hidden?: boolean): void`

Adds a new worksheet to the workbook.

**Parameters:**
- `sheetName` - Name of the worksheet (max 31 characters, invalid characters will be sanitized)
- `hidden` - Optional. If `true`, the sheet will be hidden (default: `false`)

**Example:**
```typescript
writer.addSheet('Sales Data');
writer.addSheet('Hidden Sheet', true);
```

#### `writeSheet(rows: any[][], headers?: string[] | null, doAutofilter?: boolean): void`

Writes data to the current sheet.

**Parameters:**
- `rows` - 2D array of data to write
- `headers` - Optional. Array of header strings (if not provided, first row is used as headers)
- `doAutofilter` - Optional. If `true`, adds autofilter to headers (default: `true`)

**Supported Data Types:**
- `string` - Written as shared string
- `number` - Written as number (integer or double)
- `boolean` - Written as boolean
- `Date` - Written as Excel date serial number. The serial is taken from the UTC
  time (`getTime()`), not local wall-clock time: `new Date(2024, 5, 30)` in UTC+2
  is stored as 2024-06-29 22:00 and shows that way in Excel. Use `Date.UTC(...)`
  for date-only values
- `null` / `undefined` - Written as empty cell

**Example:**
```typescript
writer.writeSheet([
    ['Product', 'Price', 'Quantity'],
    ['Widget A', 19.99, 100],
    ['Widget B', 29.99, 50]
]);

// With separate headers
writer.writeSheet(
    [['Widget A', 19.99, 100], ['Widget B', 29.99, 50]],
    ['Product', 'Price', 'Quantity']
);
```

#### `finalize(): Promise<void>`

Finalizes and closes the workbook. Must be called after all data is written.

**Example:**
```typescript
await writer.finalize();
```

---

## XlsxWriter

Writer for Excel Open XML Workbook (.xlsx) format. Has the same API as XlsbWriter.

### Constructor

```typescript
new XlsxWriter(filePath: string)
```

### Methods

All methods are identical to [XlsbWriter](#xlsbwriter):
- `addSheet(sheetName: string, hidden?: boolean): void`
- `writeSheet(rows: any[][], headers?: string[] | null, doAutofilter?: boolean): void`
- `finalize(): Promise<void>`

---

## XlsbReader

High-performance reader for Excel Binary Workbook (.xlsb) format.

### Constructor

```typescript
new XlsbReader()
```

### Properties

#### `fieldCount: number`

Number of columns in the current row.

#### `resultsCount: number`

Number of sheets in the workbook.

### Methods

#### `open(path: string, readSharedStrings?: boolean): Promise<void>`

Opens an XLSB file for reading.

**Parameters:**
- `path` - Path to the XLSB file
- `readSharedStrings` - Optional. If `true`, reads shared strings table (default: `true`)

**Example:**
```typescript
const reader = new XlsbReader();
await reader.open('data.xlsb');
```

#### `getSheetNames(): string[]`

Returns array of worksheet names.

**Example:**
```typescript
const sheets = reader.getSheetNames();
console.log(sheets); // ['Sheet1', 'Sheet2']
```

#### `read(): Promise<boolean>`

Reads the next row. Returns `true` if a row was read, `false` if end of sheet reached.
Always use `await` (both XLSB and XLSX are async in v2+).

Empty rows: gaps in the sheet (missing `<row>` / `BrtRowHdr`, including leading
gaps) are returned as empty rows (`[]`, `getValue()` → `null`). Like xlspy,
gap positions are yielded; unlike xlspy's padded `[None] * cols`, this reader
yields a zero-length row so callers must not assume `fieldCount` width.
The 1904 date system (`workbookPr date1904` / `BrtWbProp`) is honoured.

**Example:**
```typescript
while (await reader.read()) {
    // Process row
}
```

#### `selectSheet(nameOrIndex: string | number): Promise<void>`

Selects a worksheet for subsequent `read()` calls. Resets row/EOF and `fieldCount`.
Default is sheet 0.

```typescript
await reader.selectSheet('Second');
await reader.selectSheet(0);
```

#### `nextResult(): Promise<boolean>`

Advances to the next worksheet. Returns `true` if another sheet was selected,
`false` at the end.

```typescript
do {
  while (await reader.read()) { /* rows */ }
} while (await reader.nextResult());
```

#### `getValue(columnIndex: number): CellValue`

Gets value at the specified column index in the current row.

**Parameters:**
- `columnIndex` - Zero-based column index

**Returns:** `CellValue` (string, number, boolean, Date, bigint, or null/undefined)

Date cells are returned as `Date` objects built from the serial in UTC, the same
convention the writers use, so a write/read round trip keeps the same instant.

**Example:**
```typescript
while (await reader.read()) {
    const name = reader.getValue(0);
    const age = reader.getValue(1);
    console.log(`${name}: ${age}`);
}
```

---

## XlsxReader

Reader for Excel Open XML Workbook (.xlsx) format.

### Constructor

```typescript
new XlsxReader()
```

### Properties

Same as [XlsbReader](#xlsbreader):
- `fieldCount: number`
- `resultsCount: number`

### Methods

#### `open(path: string, readSharedStrings?: boolean): Promise<void>`

Opens an XLSX file for reading.

#### `close(): Promise<void>`

Closes the reader and releases resources.

#### `getSheetNames(): string[]`

Returns array of worksheet names.

#### `read(): Promise<boolean>`

Reads the next row. Returns `true` if a row was read. Always use `await`.
Empty gap rows (including leading) are returned as empty rows; see
[XlsbReader](#xlsbreader) for `selectSheet`/`nextResult`, which behave identically.

#### `selectSheet(nameOrIndex: string | number): Promise<void>`

See [XlsbReader](#xlsbreader).

#### `nextResult(): Promise<boolean>`

See [XlsbReader](#xlsbreader).

#### `getValue(columnIndex: number): CellValue`

Gets value at the specified column index.

---

## XlsxUpdater

Updates the data of a worksheet inside an existing `.xlsx` file without rebuilding
the workbook. Everything outside the target sheet's cell data — pivot tables and
their caches, other sheets, styles, themes, defined names — is preserved, which
makes it possible to refresh a data sheet that pivot tables / charts / formulas
are wired to.

### Constructor

```typescript
new XlsxUpdater(filePath: string)
```

**Parameters:**
- `filePath` - Path to an existing `.xlsx` file

Throws if the file does not exist, is not a valid XLSX package, or is an XLSB.
Macro-enabled `.xlsm` packages can use the explicit [XlsmUpdater](#xlsmupdater)
alias; unknown VBA members are preserved.

**Example:**
```typescript
const updater = new XlsxUpdater('report.xlsx');
```

### Methods

#### `getSheetNames(): string[]`

Returns the worksheet names in workbook order.

#### `replaceSheetData(sheetName: string, rows: CellValue[][], options?: ReplaceSheetDataOptions): void`

Clears the entire cell data of the target sheet and writes new rows in its place.

**Parameters:**
- `sheetName` - Name of the worksheet to replace. Throws if not found.
- `rows` - 2D array of new data (`null` / `undefined` cells stay empty).
  Trailing rows that contain no non-empty cells are trimmed so the written
  range matches the actual data (pivot tables won't pick up "(blank)" items
  from padding rows).
- `options` - Optional object:
  - `headers?: string[]` - Optional header row written into row 1
  - `styleFallback?: 'inherit' | 'general'` - Style strategy for new cells
    (default `'inherit'`)

**Behavior:**
- The sheet's `<dimension>` and `<autoFilter>` ranges are recalculated and end
  exactly at the last data row.
- When `headers` are omitted, the first data row lands in row 1 and becomes the
  pivot's field names — always pass `headers` matching the original columns.
- New strings are appended to `xl/sharedStrings.xml` (existing entries and their
  indices are never modified); files without a shared string table use inline strings.
- With `styleFallback: 'inherit'` (default), new cells reuse the dominant style
  of their column from the existing data, and `Date` cells fall back to a date
  style found in `styles.xml` when the column has none.
- Pivot tables whose `worksheetSource` references the updated sheet get their
  `ref` and `recordCount` adjusted, and `refreshOnLoad="1"` is added to the
  pivot table definitions so Excel refreshes them on open.
- The stale calculation chain (`xl/calcChain.xml`) is removed with its
  `Content_Types`/rels entries; Excel rebuilds it on open.
- ListObject tables (`xl/tables/*.xml`) anchored at `A1` get their `<table>`
  and `<autoFilter>` `ref` updated to the new dimension. Tables whose column
  count changes are left untouched (their `tableColumns` would need rebuilding),
  and ranges are capped so they never overlap another table.

**Example:**
```typescript
const updater = new XlsxUpdater('report.xlsx');
updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME'] });
updater.save('report_new.xlsx');
```

#### `save(outputPath?: string): void`

Writes the updated workbook to disk. When `outputPath` is omitted, the source
file is overwritten in place.

#### `replaceSheetDataStream(sheetName: string, rows: RowSource, options?: ReplaceSheetDataOptions): Promise<void>`

Replaces worksheet rows from a synchronous or asynchronous one-pass row source.
The row source is processed incrementally and is not collected into a complete
2D array. Use `saveStreaming()` to also avoid materialising the final ZIP buffer.

#### `saveStreaming(outputPath?: string): Promise<void>`

Writes the updated package through a temporary file and installs it atomically.
Staged worksheet parts are streamed; `save()` remains the simpler synchronous
buffering path.

#### `toBuffer(): Buffer`

Returns the updated workbook as an in-memory ZIP buffer.

## XlsmUpdater

`XlsmUpdater` is an explicit alias of `XlsxUpdater` for macro-enabled `.xlsm`
packages. It has the same API as `XlsxUpdater` and preserves VBA parts without
parsing or modifying them.

---

## XlsbUpdater

Updates the data of a worksheet inside an existing `.xlsb` file without
rebuilding the workbook. Everything outside the target sheet's rows — pivot
tables and their caches, other sheets, styles, column widths, views — is
preserved byte-for-byte.

The API is identical to [XlsxUpdater](#xlsxupdater); see that section for the
behavior of `replaceSheetData` (inherited column styles, shared-strings
appending, pivot cache range + `refreshOnLoad`).

### Constructor

```typescript
new XlsbUpdater(filePath: string)
```

**Parameters:**
- `filePath` - Path to an existing `.xlsb` file

Throws if the file does not exist, is not a valid XLSB, or is an XLSX
(XLSB-only updater).

### Methods

#### `getSheetNames(): string[]`

Returns the worksheet names in workbook order.

#### `replaceSheetData(sheetName: string, rows: CellValue[][], options?: ReplaceSheetDataOptions): void`

Clears the entire cell data of the target sheet and writes new rows in its place.

**Parameters:**
- `sheetName` - Name of the worksheet to replace. Throws if not found.
- `rows` - 2D array of new data (`null` / `undefined` cells stay empty).
  Trailing rows that contain no non-empty cells are trimmed so the written
  range matches the actual data (pivot tables won't pick up "(blank)" items
  from padding rows).
- `options` - Optional object:
  - `headers?: string[]` - Optional header row written into row 1
  - `styleFallback?: 'inherit' | 'general'` - Style strategy for new cells
    (default `'inherit'`)

**Behavior:**
- Only the worksheet's row records are replaced; column widths, views, and all
  other sheet records are preserved. Auto-filter ranges are recalculated.
- The sheet's dimension record and the pivot cache source range end exactly at
  the last data row.
- When `headers` are omitted, the first data row lands in row 1 and becomes the
  pivot's field names — always pass `headers` matching the original columns.
- New strings are appended to `xl/sharedStrings.bin` (existing entries and their
  indices are never modified).
- With `styleFallback: 'inherit'` (default), new cells reuse the dominant cell
  style (xf) of their column from the existing data, and `Date` cells fall back
  to a date style found in `styles.bin` when the column has none.
- Pivot caches whose `worksheetSource` references the updated sheet get their
  range and record count adjusted, and the `refreshOnLoad` flag is set so Excel
  refreshes them on open.
- The stale calculation chain (`xl/calcChain.bin`) is removed with its
  `Content_Types`/rels entries.
- XLSB ListObject tables are **not** patched (documented limitation); only the
  sheet dimension (`BrtWsDim`) and pivot caches are updated.

**Example:**
```typescript
const updater = new XlsbUpdater('report.xlsb');
updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME'] });
updater.save('report_new.xlsb');
```

#### `save(outputPath?: string): void`

Writes the updated workbook to disk. When `outputPath` is omitted, the source
file is overwritten in place.

#### `toBuffer(): Buffer`

Returns the updated workbook as an in-memory ZIP buffer.

---

## Excel conversion

Excel conversion uses the installed desktop Microsoft Excel through COM. It is
available in the Node.js build on Windows only; the browser build does not
expose these functions.

### `ExcelConversionOptions`

```typescript
interface ExcelConversionOptions {
    overwrite?: boolean;
}
```

`overwrite` defaults to `false`. Existing destination files are rejected unless
it is explicitly set to `true`.

### `convertXlsbToXlsx`

```typescript
convertXlsbToXlsx(
    inputPath: string,
    outputPath: string,
    options?: ExcelConversionOptions
): Promise<void>
```

Converts an existing `.xlsb` file to `.xlsx` using Excel `SaveAs` format 51.

### `convertXlsxToXlsb`

```typescript
convertXlsxToXlsb(
    inputPath: string,
    outputPath: string,
    options?: ExcelConversionOptions
): Promise<void>
```

Converts an existing `.xlsx` file to `.xlsb` using Excel `SaveAs` format 50.

Both functions reject invalid extensions, missing files, protected
destinations, unavailable COM automation, and Excel errors. The source is
opened read-only and the destination is first written to a temporary file.

```typescript
await convertXlsbToXlsx('report.xlsb', 'report.xlsx');
await convertXlsxToXlsb('report.xlsx', 'report.xlsb', { overwrite: true });
```

## ReaderFactory

Factory class for creating appropriate reader based on file extension.

### Methods

#### `static create(filePath: string): XlsbReader | XlsxReader`

Creates a reader instance based on file extension.

**Parameters:**
- `filePath` - Path to the Excel file

**Returns:** `XlsbReader` for .xlsb files, `XlsxReader` for .xlsx files

**Example:**
```typescript
const reader = ReaderFactory.create('data.xlsb');
await reader.open('data.xlsb');
```

---

## Types

### Excel limits (throw, xlspy parity)

```typescript
import { MAX_ROWS, MAX_COLS, MAX_CELL_CHARS } from '@justybase/spreadsheet-tasks';
// MAX_ROWS = 1_048_576, MAX_COLS = 16_384, MAX_CELL_CHARS = 32_767 UTF-16 units
```

Writers throw when a limit is exceeded (no truncation, workbook not published).
`writeSheet` also throws when a row/header is longer than `columnCount`, matching
streaming `writeRow` behaviour. Astral characters count as 2 units
(`'😀'.length === 2`).

### Supported Cell Types

| TypeScript Type | Excel Cell Type |
|-----------------|-----------------|
| `string` | Shared String |
| `number` (integer) | RK Number |
| `number` (float) | Double |
| `boolean` | Boolean |
| `Date` | Date Serial Number |
| `null` / `undefined` | Empty Cell |

### `RowSource`

```typescript
type RowSource =
  | Iterable<ReadonlyArray<CellValue>>
  | AsyncIterable<ReadonlyArray<CellValue>>;
```

Used by `replaceSheetDataStream()` in both updater classes.

### Sheet Name Restrictions

Sheet names have the following restrictions:
- Maximum 31 characters
- Cannot contain: `\ / * ? [ ] :`
- Cannot be empty

Invalid characters are automatically removed when adding sheets.
