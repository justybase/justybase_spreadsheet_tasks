/**
 * Deterministic datasets for the I/O benchmark.
 *
 * The layouts mirror xlspy's `benchmarks/bench_io.py` where it matters: a
 * `mixed` set cycling through text, ints, floats, dates and booleans, plus
 * text-heavy and number-heavy sets. Nothing here is random, so two runs on
 * the same machine compare like with like.
 */

export type DatasetName = 'mixed' | 'strings' | 'numbers' | 'dates' | 'sparse';

export interface DatasetSpec {
    name: DatasetName;
    rows: number;
    cols: number;
}

export const BASE_DATE_UTC = Date.UTC(2023, 0, 1);

/** A deterministic, reasonably wide string: 1000 distinct prefixes. */
function textValue(row: number, col: number): string {
    return `String_${row % 1000}_${col}`;
}

function sparseValue(row: number, col: number): string | number | null {
    if ((row + col) % 7 === 0) return null;
    if (col % 3 === 0) return row * 1000 + col;
    return `S_${row % 97}_${col}`;
}

/** Build one row of the requested dataset. */
export function makeRow(dataset: DatasetName, row: number, cols: number, baseDate = BASE_DATE_UTC): unknown[] {
    const out = new Array<unknown>(cols);
    for (let col = 0; col < cols; col++) {
        switch (dataset) {
            case 'mixed': {
                const kind = col % 5;
                if (kind === 0) out[col] = textValue(row, col);
                else if (kind === 1) out[col] = row * (col + 1);
                else if (kind === 2) out[col] = row / (col + 1);
                else if (kind === 3) out[col] = new Date(baseDate + (row % 3000) * 86400000);
                else out[col] = (row + col) % 2 === 0;
                break;
            }
            case 'strings':
                out[col] = textValue(row, col);
                break;
            case 'numbers':
                out[col] = row * (col + 1) + col / 10;
                break;
            case 'dates':
                out[col] = new Date(baseDate + ((row * 7 + col) % 3000) * 86400000);
                break;
            case 'sparse':
                out[col] = sparseValue(row, col);
                break;
        }
    }
    return out;
}

export function makeRows(dataset: DatasetName, rows: number, cols: number): unknown[][] {
    const out: unknown[][] = new Array(rows);
    for (let row = 0; row < rows; row++) out[row] = makeRow(dataset, row, cols);
    return out;
}

export function makeHeaders(cols: number): string[] {
    return Array.from({ length: cols }, (_, col) => `Header_${col + 1}`);
}

/** Number of non-null cells the dataset produces; a cheap read-correctness oracle. */
export function expectedNonNullCount(dataset: DatasetName, rows: number, cols: number): number {
    let count = 0;
    for (let row = 0; row < rows; row++) {
        for (const value of makeRow(dataset, row, cols)) {
            if (value !== null && value !== undefined) count++;
        }
    }
    return count;
}

export interface BenchCase {
    id: string;
    kind: 'write' | 'read';
    format: 'xlsx' | 'xlsb';
    mode: 'batch' | 'streaming';
    dataset: DatasetName;
    rows: number;
    cols: number;
    sharedStrings?: boolean;
    compressionLevel?: number;
    notes?: string;
}

export function defaultCases(quick: boolean): BenchCase[] {
    const mixed = quick
        ? { rows: 5_000, cols: 20 }
        : { rows: 50_000, cols: 50 };
    const numbers = quick
        ? { rows: 20_000, cols: 10 }
        : { rows: 1_000_000, cols: 10 };
    const strings = quick
        ? { rows: 10_000, cols: 10 }
        : { rows: 200_000, cols: 20 };

    const cases: BenchCase[] = [];
    for (const format of ['xlsx', 'xlsb'] as const) {
        cases.push({
            id: `${format}-write-mixed-batch`,
            kind: 'write', format, mode: 'batch', dataset: 'mixed', ...mixed,
        });
        cases.push({
            id: `${format}-write-mixed-stream`,
            kind: 'write', format, mode: 'streaming', dataset: 'mixed', ...mixed,
        });
        cases.push({
            id: `${format}-write-numbers-batch`,
            kind: 'write', format, mode: 'batch', dataset: 'numbers', ...numbers,
        });
        cases.push({
            id: `${format}-write-strings-inline`,
            kind: 'write', format, mode: 'streaming', dataset: 'strings', ...strings,
            sharedStrings: false,
            notes: 'inline strings; falls back to shared when not supported',
        });
        cases.push({
            id: `${format}-write-strings-shared`,
            kind: 'write', format, mode: 'streaming', dataset: 'strings', ...strings,
            sharedStrings: true,
        });
        cases.push({
            id: `${format}-read-mixed`,
            kind: 'read', format, mode: 'batch', dataset: 'mixed', ...mixed,
        });
        cases.push({
            id: `${format}-read-numbers`,
            kind: 'read', format, mode: 'batch', dataset: 'numbers', ...numbers,
        });
    }
    return cases;
}
