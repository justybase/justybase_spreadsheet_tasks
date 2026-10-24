# @justybase/spreadsheet-tasks

High-performance TypeScript library for reading and writing Excel files in XLSB and XLSX formats.

[![npm version](https://img.shields.io/npm/v/@justybase/spreadsheet-tasks.svg)](https://www.npmjs.com/package/@justybase/spreadsheet-tasks)
[![CI](https://github.com/justybase/justybase_spreadsheet_tasks/actions/workflows/ci.yml/badge.svg)](https://github.com/justybase/justybase_spreadsheet_tasks/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Table of Contents

- [Features](#features)
- [Benchmark Results](#benchmark-results)
- [Requirements](#requirements)
- [Installation](#installation)
- [Browser Support](#browser-support)
- [Quick Start](#quick-start)
- [API Documentation](#api-documentation)
- [Examples](#examples)
- [When to Use XLSB vs XLSX](#when-to-use-xlsb-vs-xlsx)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Features

- **High Performance** — XLSB is about 3.3× faster to write and 2.3× faster to read than XLSX (see [benchmarks](#benchmark-results))
- **Small File Size** — XLSB files are ~47% smaller than equivalent XLSX files
- **TypeScript First** — Full TypeScript support with type definitions
- **Dual Format Support** — Read and write both XLSB and XLSX formats
- **Lightweight Node Dependencies** — Small ZIP-focused runtime deps (`adm-zip`, `archiver`, `yauzl`); no heavy Excel framework
- **Streaming Support** — Efficient memory usage for large files
- **Multiple Sheets** — Support for multiple worksheets per workbook
- **Auto-filter** — Automatic filter headers support
- **Excel COM Conversion** — Convert XLSB and XLSX through desktop Microsoft Excel on Windows
- **In-place Updates** — Replace the data of a worksheet inside an existing XLSX, XLSM or XLSB while keeping pivot tables, other sheets and formatting untouched

## Benchmark Results

Measured with `npm run benchmark` (50K rows). Your results may vary by hardware and dataset.

| Operation | XLSB | XLSX | Performance Gain |
|-----------|------|------|------------------|
| Write (50K rows) | 140 ms | 467 ms | **3.3× faster** |
| Read (50K rows) | 118 ms | 276 ms | **2.3× faster** |
| File Size | 1.49 MB | 2.83 MB | **47% smaller** |

## Requirements

- **Node.js** 16 or newer ([`engines`](./package.json) in `package.json`)
- XLSB/XLSX conversion additionally requires Windows, PowerShell and desktop Microsoft Excel with COM automation

## Installation

```bash
npm install @justybase/spreadsheet-tasks
```

CommonJS:

```js
const { XlsbWriter } = require('@justybase/spreadsheet-tasks');
```

ES modules:

```js
import { XlsbWriter } from '@justybase/spreadsheet-tasks';
```

## Browser Support

Browser writers live in TypeScript under `src/browser/` and are bundled with:

```bash
npm run build:browser
```

This produces:

- `browser/justybase-spreadsheet.min.js` — IIFE (`JustybaseSpreadsheet` global)
- `browser/browser-spreadsheet.js` — ESM for local demos (`demo.html`, `test.html`)

The browser bundle has **no runtime npm dependencies** (ZIP logic is inlined).

```html
<script type="module">
  import { downloadXlsb, downloadXlsx, F } from './browser-spreadsheet.js';

  const headers = ['ID', 'Name', 'Score'];
  const rows = [
    [1, 'Alice', 99.5],
    [2, 'Bob', 88.0],
  ];

  document.getElementById('downloadBtn').addEventListener('click', () => {
    downloadXlsb('report.xlsb', rows, headers);
  });
</script>
```

**Features included in the browser build:**

- Built-in ZIP archiving (no external libraries in the bundle)
- Memory-efficient `Uint8Array` buffer handling
- Auto-fitted column widths
- Frozen and styled headers (styles: `bold`, `fill`, `bold+fill`)
- Auto-filters enabled by default

See [CHANGELOG.md](./CHANGELOG.md) for v2 breaking changes (`await reader.read()`).

## Quick Start

### Writing Excel Files

```typescript
import { XlsbWriter, XlsxWriter } from '@justybase/spreadsheet-tasks';

// Create XLSB file (recommended for performance)
const xlsbWriter = new XlsbWriter('output.xlsb');
xlsbWriter.addSheet('Sheet1');
xlsbWriter.writeSheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago']
]);
await xlsbWriter.finalize();

// Or create XLSX file for compatibility
const xlsxWriter = new XlsxWriter('output.xlsx');
xlsxWriter.addSheet('Sheet1');
xlsxWriter.writeSheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago']
]);
await xlsxWriter.finalize();
```

### Streaming API (for Large Datasets)

For large datasets that don't fit in memory, use the streaming API to write rows one at a time:

```typescript
import { XlsbWriter } from '@justybase/spreadsheet-tasks';

const writer = new XlsbWriter('large-output.xlsb');

// Start a sheet with column count and optional headers
writer.startSheet('Data', 5, ['ID', 'Name', 'Value', 'Date', 'Active']);

// Write rows one at a time - no need to load all data in memory
for (let i = 0; i < 1_000_000; i++) {
    writer.writeRow([
        i + 1,
        `User_${i + 1}`,
        Math.random() * 10000,
        new Date(),
        i % 2 === 0
    ]);
}

// Finalize the sheet
writer.endSheet();

// You can create multiple sheets
writer.startSheet('MoreData', 3, ['Col1', 'Col2', 'Col3']);
// ... write more rows
writer.endSheet();

await writer.finalize();
```

**Benefits of Streaming API:**

- Constant memory usage regardless of dataset size
- Write millions of rows without loading all data into RAM
- Generate data on-the-fly from databases, APIs, or other sources
- Same performance as batch mode

### Compression Level

XLSX and XLSB writers compress the package with zlib. The default is level 6; pass a
level from 0 (store) to 9 (smallest, slowest) to trade size for speed:

```typescript
const writer = new XlsxWriter('fast.xlsx', { compressionLevel: 1 });
```

Out-of-range values throw a `RangeError`.

### Reading Excel Files

```typescript
import { XlsbReader, XlsxReader, ReaderFactory } from '@justybase/spreadsheet-tasks';

// Using ReaderFactory (auto-detects format)
const reader = ReaderFactory.create('data.xlsb');
await reader.open('data.xlsb');

console.log('Sheet names:', reader.getSheetNames());

while (await reader.read()) {
    const row = [];
    for (let i = 0; i < reader.fieldCount; i++) {
        row.push(reader.getValue(i));
    }
    console.log(row);
}

// Or use specific reader
const xlsbReader = new XlsbReader();
await xlsbReader.open('data.xlsb');
// ... same API as above
```

### Updating Data in an Existing Workbook

`XlsxUpdater` and `XlsbUpdater` replace the cell data of a worksheet inside an
existing workbook without rebuilding it. Everything else — pivot tables, charts,
other sheets, styles, defined names — is preserved, so you can keep an advanced
Excel file (like a report with pivot tables wired to a data sheet) and only
refresh its data:

```typescript
import { XlsxUpdater, XlsbUpdater, XlsmUpdater } from '@justybase/spreadsheet-tasks';

// XLSX
const xlsx = new XlsxUpdater('report.xlsx');
xlsx.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
xlsx.save();                 // overwrite in place, or xlsx.save('report_new.xlsx')

// XLSB — the same API
const xlsb = new XlsbUpdater('report.xlsb');
xlsb.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
xlsb.save('report_new.xlsb');

// XLSM uses the XLSX updater and keeps opaque VBA parts intact.
const xlsm = new XlsmUpdater('report.xlsm');
xlsm.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
xlsm.save('report_new.xlsm');
```

For a result set that should not be collected into a 2D array, use the additive
streaming API. It accepts either a synchronous iterable or an async iterable and
writes the final ZIP through a temporary file before installing it atomically:

```typescript
const updater = new XlsxUpdater('large-report.xlsx');
await updater.replaceSheetDataStream('data1', databaseRows(), { headers });
await updater.saveStreaming('large-report-updated.xlsx');
```

`save()` remains available for the existing synchronous API. Both save methods
replace their destination atomically after a successful write.

**What happens under the hood:**

- The target sheet's cell data is cleared and replaced with the new rows.
- New strings are appended to the shared strings table (existing entries and
  their indices are never moved, so other sheets stay valid).
- New cells reuse the dominant style of their column from the original data
  (dates stay dates, numbers stay numbers). Pass `styleFallback: 'general'`
  to write everything as General.
- Pivot tables whose cache source is the updated sheet get their source range
  and record count adjusted, and `refreshOnLoad` is enabled so Excel refreshes
  them automatically on open.

> **Tip:** always pass the `headers` matching the original columns. When
> `headers` are omitted, the first data row lands in row 1 and becomes the
> pivot's field names (Excel reads the first row of the source range as
> headers), so the pivot's columns would be replaced by the first record's
> values.

### Converting XLSB and XLSX with Excel

On Windows, existing workbooks can be converted by asking the installed desktop
Microsoft Excel to perform `SaveAs` through COM. The source file is opened
read-only and the destination is written separately:

```typescript
import {
    convertXlsbToXlsx,
    convertXlsxToXlsb,
} from '@justybase/spreadsheet-tasks';

await convertXlsbToXlsx('report.xlsb', 'report.xlsx');
await convertXlsxToXlsb('report.xlsx', 'report.xlsb');

// Existing destinations are protected by default.
await convertXlsbToXlsx('report.xlsb', 'report.xlsx', { overwrite: true });
```

Conversion is a Node.js-only API and requires Windows plus desktop Microsoft
Excel. Unsupported files, unavailable COM automation, and Excel conversion
errors are reported as rejected promises. VBA/macros are not guaranteed when
the target format does not support them.

## API Documentation

See [API Documentation](./docs/API.md) for detailed API reference.

## Examples

Check out the [examples](./examples) folder for more usage examples:

- [Basic Write](./examples/basic-write.ts) — Writing data to Excel files
- [Basic Read](./examples/basic-read.ts) — Reading data from Excel files
- [Multiple Sheets](./examples/multiple-sheets.ts) — Working with multiple worksheets
- [Update Existing Workbook](./examples/update-existing.ts) — Replacing the data of a sheet in an existing XLSX/XLSB with pivot tables
- [Large Dataset](./examples/large-dataset.ts) — Handling large datasets efficiently
- [Streaming Example](./examples/streaming-example.ts) — Write millions of rows without loading all data in memory

Run examples with:

```bash
npm run build
npx ts-node examples/basic-write.ts
```

## When to Use XLSB vs XLSX

| Use XLSB When | Use XLSX When |
|---------------|---------------|
| Performance is critical | Need maximum compatibility |
| Working with large datasets | Sharing with older Excel versions |
| Internal/backend processing | Human-readable XML is needed |
| Storage space is limited | Third-party integrations require it |

## Contributing

Contributions are welcome. To work on the project locally:

```bash
git clone https://github.com/justybase/justybase_spreadsheet_tasks.git
cd justybase_spreadsheet_tasks
npm install
npm run build
npm test
```

Before opening a pull request:

1. Run `npm run lint` and `npm test`.
2. Add or update tests if you change behavior.
3. Update [CHANGELOG.md](./CHANGELOG.md) for user-visible changes.

Please open an issue or pull request on [GitHub](https://github.com/justybase/justybase_spreadsheet_tasks).

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Inspired by and reimplemented from [SpreadSheetTasks](https://github.com/justybae/SpreadSheetTasks)
