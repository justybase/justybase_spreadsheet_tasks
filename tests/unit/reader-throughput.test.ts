/**
 * Reader throughput guard.
 *
 * The row parser must stay O(cell), not O(cell x sheet size): an unbounded
 * `indexOf` for a tag that is absent from the cell used to scan the whole
 * remaining sheet for every cell. A 20k x 10 sheet then took minutes instead
 * of milliseconds. The bound below is deliberately loose (seconds, not
 * milliseconds) so it only fails on that class of regression.
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { CellValue } from '../../src/Formats';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsxWriter } from '../../src/XlsxWriter';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-reader-throughput-'));
const MAX_READ_MS = 10_000;

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

async function writeSheet(
    file: string,
    rows: CellValue[][],
    options: { useSharedStrings?: boolean } = {},
): Promise<void> {
    const writer = new XlsxWriter(file, {
        compressionLevel: 0,
        useSharedStrings: options.useSharedStrings,
    });
    writer.addSheet('Data');
    writer.writeSheet(rows, null, { autofitRows: 0 });
    await writer.finalize();
}

async function readAllRows(file: string): Promise<{ rows: number; first: unknown[]; last: unknown[]; ms: number }> {
    const reader = new XlsxReader();
    await reader.open(file);
    await reader.selectSheet(0);
    const startedAt = Date.now();
    let rows = 0;
    let first: unknown[] = [];
    let last: unknown[] = [];
    while (true) {
        const row = reader.readRow();
        if (row === null) break;
        if (rows === 0) first = row.slice();
        last = row.slice();
        rows++;
    }
    const ms = Date.now() - startedAt;
    await reader.close();
    return { rows, first, last, ms };
}

async function testNumericSheet(): Promise<void> {
    const rowCount = 20_000;
    const colCount = 10;
    const file = path.join(outputDir, 'numbers.xlsx');
    const rows = Array.from({ length: rowCount }, (_, row) =>
        Array.from({ length: colCount }, (_, col) => row * (col + 1) + col / 10),
    );
    await writeSheet(file, rows);

    const result = await readAllRows(file);
    assert(result.rows === rowCount, `expected ${rowCount} rows, got ${result.rows}`);
    assert(result.ms < MAX_READ_MS, `reading ${rowCount * colCount} numbers took ${result.ms} ms`);
    assert(result.first[0] === 0 && result.first[9] === 0.9,
        `first row values, got ${JSON.stringify(result.first)}`);
    assert(result.last[0] === (rowCount - 1) && result.last[9] === (rowCount - 1) * 10 + 0.9,
        `last row values, got ${JSON.stringify(result.last)}`);
}

async function testInlineStrings(): Promise<void> {
    const rowCount = 10_000;
    const file = path.join(outputDir, 'inline.xlsx');
    const rows = Array.from({ length: rowCount }, (_, row) => [`text_${row % 500}_a`, row]);
    // Inline cells have no <v>, so this covers the other unbounded-lookup path.
    await writeSheet(file, rows, { useSharedStrings: false });

    const result = await readAllRows(file);
    assert(result.rows === rowCount, `expected ${rowCount} inline rows, got ${result.rows}`);
    assert(result.ms < MAX_READ_MS, `reading ${rowCount} inline strings took ${result.ms} ms`);
    assert(result.first[0] === 'text_0_a' && result.last[0] === `text_${(rowCount - 1) % 500}_a`,
        'inline string values round-trip');
}

async function testNumericPrecision(): Promise<void> {
    // Integers and <=15 significant digits use the allocation-free fast path;
    // everything else must fall back to parseFloat without changing the value.
    const values = [
        0, 0.1, 0.07, 2.2, -1.5, 123456789012345, 999999999999999,
        1.2345678901234567, 0.0000123456789012345, 1234567890.12345,
        1e-7, 1e21, 1e15 + 1, 5e-324, 1.7976931348623157e308,
    ];
    const file = path.join(outputDir, 'precision.xlsx');
    await writeSheet(file, values.map(value => [value]));

    const reader = new XlsxReader();
    await reader.open(file);
    await reader.selectSheet(0);
    for (let i = 0; i < values.length; i++) {
        const row = reader.readRow();
        assert(row !== null, `missing row ${i}`);
        assert(Object.is(row![0], values[i]),
            `value ${i}: expected ${values[i]}, got ${String(row![0])}`);
    }
    assert(reader.readRow() === null, 'no extra rows');
    await reader.close();
}

async function main(): Promise<void> {
    await testNumericSheet();
    await testInlineStrings();
    await testNumericPrecision();
    console.log('Reader throughput tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
