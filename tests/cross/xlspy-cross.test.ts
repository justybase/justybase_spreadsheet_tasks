/**
 * Cross-validation suite between this TypeScript library and xlspy (Python).
 *
 * Two directions are covered:
 *   - TS writes -> xlspy reads (via tests/cross/xlspy_bridge.py)
 *   - xlspy writes -> TS reads
 * plus updater round-trips in both directions and the committed xlspy golden
 * fixtures (openpyxl/xlsxwriter-authored, incl. the 1904 date system).
 *
 * Requires Python with xlspy importable. If xlspy is unavailable the suite
 * prints SKIP and exits 0 (same pattern as the optional pivot fixtures).
 * The golden-fixture tests additionally need the xlspy repo checkout; set
 * XLSPY_REPO or place it next to this repo (../xlspy). Otherwise those tests
 * skip while the bridge-generated round-trips still run.
 */
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxUpdater } from '../../src/XlsxUpdater';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import type { CellValue } from '../../src/Formats';

const BRIDGE = path.join(__dirname, 'xlspy_bridge.py');
const PYTHON = process.env.XLSPY_PYTHON ?? 'python3';
const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'xlspy-cross-'));

type BridgeCell = string | number | boolean | null | { datetime: string };
type BridgeRow = BridgeCell[];

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

// ---------------------------------------------------------------------------
// Bridge plumbing
// ---------------------------------------------------------------------------

function bridge(args: string[], job: unknown): any {
    const res = spawnSync(PYTHON, [BRIDGE, ...args], {
        input: JSON.stringify(job),
        encoding: 'utf8',
        maxBuffer: 64 * 1024 * 1024,
    });
    if (res.status !== 0) {
        throw new Error(`bridge ${args.join(' ')} failed: ${(res.stderr ?? '').slice(0, 2000)}`);
    }
    return JSON.parse(res.stdout || '{}');
}

function xlspyAvailable(): boolean {
    try {
        const res = spawnSync(PYTHON, [BRIDGE, 'check'], {
            input: '{}', encoding: 'utf8', timeout: 60000,
        });
        if (res.status !== 0) return false;
        return (JSON.parse(res.stdout || '{}') as any).ok === true;
    } catch {
        return false;
    }
}

function xlspyRepoDir(): string | null {
    const cand = process.env.XLSPY_REPO
        ?? path.resolve(__dirname, '..', '..', '..', 'xlspy');
    const data = path.join(cand, 'tests', 'data');
    return fs.existsSync(data) ? data : null;
}

// ---------------------------------------------------------------------------
// Normalization: both readers reduce to the same JSON shape.
// ---------------------------------------------------------------------------

/** Date -> 'd:YYYY-MM-DDTHH:mm:ss'; xlspy {"datetime": iso} -> same form. */
function normCell(c: unknown): unknown {
    if (c instanceof Date) return `d:${c.toISOString().slice(0, 19)}`;
    if (typeof c === 'object' && c !== null && typeof (c as any).datetime === 'string') {
        const iso = (c as any).datetime as string;
        return `d:${iso.length === 10 ? `${iso}T00:00:00` : iso.slice(0, 19)}`;
    }
    if (typeof c === 'number') return Object.is(c, -0) ? 0 : c;
    if (c === undefined) return null;
    return c;
}

/** Strip trailing nulls (xlspy pads gap rows to [None]*cols; TS pads to fieldCount). */
function normRows(rows: unknown[][]): unknown[][] {
    return rows.map((r) => {
        const out = (r as unknown[]).map(normCell);
        while (out.length > 0 && out[out.length - 1] === null) out.pop();
        return out;
    });
}

function assertRowsEqual(actual: unknown[][], expected: unknown[][], label: string): void {
    const a = JSON.stringify(normRows(actual));
    const e = JSON.stringify(normRows(expected));
    assert(a === e, `${label} mismatch:\n  actual:   ${a}\n  expected: ${e}`);
}

/** Bridge {"datetime": iso} -> TS Date (UTC midnight for date-only). */
function toTsRows(rows: BridgeRow[]): CellValue[][] {
    return rows.map((r) => r.map((c): CellValue => {
        if (typeof c === 'object' && c !== null && 'datetime' in c) {
            const iso = (c as { datetime: string }).datetime;
            return new Date(iso.length === 10 ? `${iso}T00:00:00Z` : `${iso}Z`);
        }
        return c as CellValue;
    }));
}

// ---------------------------------------------------------------------------
// Shared fixture matrix (bridge encoding; TS side converts via toTsRows)
// ---------------------------------------------------------------------------

const HEADER = ['text', 'int', 'float', 'bool', 'date', 'datetime', 'hole', 'tail'];

const MATRIX: BridgeRow[] = [
    ['plain', 0, 0.0, true, { datetime: '2024-01-15' }, { datetime: '2024-01-15T10:30:00' }, null, 'end'],
    ['  padded  ', -1, -0.001, false, { datetime: '1900-01-01' }, { datetime: '1999-09-09T09:09:09' }, null, 'end'],
    ['<tag> & "q" \'x\'', 2147483647, 1e10, true, { datetime: '2173-10-14' }, { datetime: '2024-02-29T23:59:59' }, 'mid', 'end'],
    ['zażółć ⚡ 日本語 🎉', -2147483648, 1.23456789e-8, false, { datetime: '2000-02-29' }, null, null, 'end'],
    ['line\nbreak\ttab', 123456789, 0.1 + 0.2, false, { datetime: '1970-01-01' }, null, null, ''],
    [],
    [null, null, 'tail-only'],
];

const MATRIX_WITH_HEADER: BridgeRow[] = [HEADER, ...MATRIX];
const UNI_SHEET = 'Łąka ⚡_データ';
const UNI_ROWS: BridgeRow[] = [['H'], ['uni-żółć'], [123]];

// ---------------------------------------------------------------------------
// TS-side helpers
// ---------------------------------------------------------------------------

type TsReader = XlsxReader | XlsbReader;

async function readAllSheets(reader: TsReader, file: string): Promise<Record<string, unknown[][]>> {
    await reader.open(file);
    try {
        const out: Record<string, unknown[][]> = {};
        let current: unknown[][] = [];
        let name = '';
        const flush = () => { if (name) out[name] = current; };
        // First read() auto-selects sheet 0.
        let has = await reader.read();
        name = reader.actualSheetName;
        while (true) {
            while (has) {
                const row: unknown[] = [];
                for (let i = 0; i < reader.fieldCount; i++) row.push(reader.getValue(i));
                current.push(row);
                has = await reader.read();
            }
            flush();
            if (!(await reader.nextResult())) break;
            name = reader.actualSheetName;
            current = [];
            has = await reader.read();
        }
        return out;
    } finally {
        await reader.close();
    }
}

// ---------------------------------------------------------------------------
// A. Golden fixtures (TS reads files authored by openpyxl/xlsxwriter)
// ---------------------------------------------------------------------------

async function testGoldenFixtures(dataDir: string): Promise<void> {
    for (const f of ['golden_openpyxl.xlsx', 'golden_openpyxl_1904.xlsx', 'golden_xlsxwriter.xlsx']) {
        const sidecar = JSON.parse(fs.readFileSync(path.join(dataDir, `${f}.json`), 'utf8'));
        const sheets = await readAllSheets(new XlsxReader(), path.join(dataDir, f));
        for (const name of Object.keys(sidecar.sheets)) {
            assert(name in sheets, `golden ${f}: sheet ${name} missing`);
            assertRowsEqual(sheets[name], sidecar.sheets[name] as unknown[][], `golden ${f} sheet ${name}`);
        }
        console.log(`golden ok: ${f}`);
    }
}

// ---------------------------------------------------------------------------
// B. TS writes -> xlspy reads
// ---------------------------------------------------------------------------

async function testTsWriteXlspyRead(format: 'xlsx' | 'xlsb'): Promise<void> {
    const file = path.join(outputDir, `ts-write.${format}`);
    const W = format === 'xlsx' ? XlsxWriter : XlsbWriter;
    const w = new W(file);
    w.addSheet('Data');
    w.writeSheet(toTsRows(MATRIX), [...HEADER], false);
    w.addSheet('empty');
    w.writeSheet([], null, false);
    w.addSheet(UNI_SHEET);
    w.writeSheet(toTsRows(UNI_ROWS), null, false);
    await w.finalize();

    const res = bridge(['read', '--file', file], {});
    const sheets = res.sheets as Record<string, BridgeRow[]>;
    assertRowsEqual(sheets['Data'], MATRIX_WITH_HEADER, `TS->xlspy ${format} Data`);
    assertRowsEqual(sheets['empty'] ?? [], [], `TS->xlspy ${format} empty`);
    assertRowsEqual(sheets[UNI_SHEET], UNI_ROWS, `TS->xlspy ${format} unicode sheet`);
    console.log(`TS write -> xlspy read ok: ${format}`);
}

async function testTsStreamingWriteXlspyRead(format: 'xlsx' | 'xlsb'): Promise<void> {
    const file = path.join(outputDir, `ts-stream.${format}`);
    const W = format === 'xlsx' ? XlsxWriter : XlsbWriter;
    const w = new W(file);
    w.startSheet('S', 8);
    for (const r of toTsRows(MATRIX_WITH_HEADER)) {
        // Streaming writeRow requires exact column width; pad short rows.
        const padded = [...r];
        while (padded.length < 8) padded.push(null);
        w.writeRow(padded);
    }
    w.endSheet();
    await w.finalize();

    const res = bridge(['read', '--file', file], {});
    assertRowsEqual((res.sheets as any)['S'], MATRIX_WITH_HEADER, `TS stream->xlspy ${format}`);
    console.log(`TS streaming write -> xlspy read ok: ${format}`);
}

// ---------------------------------------------------------------------------
// C. xlspy writes -> TS reads
// ---------------------------------------------------------------------------

async function testXlspyWriteTsRead(format: 'xlsx' | 'xlsb'): Promise<void> {
    const file = path.join(outputDir, `py-write.${format}`);
    bridge(['write', '--format', format, '--file', file], {
        sheets: [
            { name: 'Data', rows: MATRIX_WITH_HEADER },
            { name: 'empty', rows: [] },
            { name: UNI_SHEET, rows: UNI_ROWS },
        ],
    });
    const R = format === 'xlsx' ? XlsxReader : XlsbReader;
    const sheets = await readAllSheets(new R(), file);
    assertRowsEqual(sheets['Data'], MATRIX_WITH_HEADER, `xlspy->TS ${format} Data`);
    assertRowsEqual(sheets['empty'] ?? [], [], `xlspy->TS ${format} empty`);
    assertRowsEqual(sheets[UNI_SHEET], UNI_ROWS, `xlspy->TS ${format} unicode sheet`);

    // Empty-sheet selection must not throw (regression guard on xlspy bytes).
    const r = new R();
    await r.open(file);
    try {
        await r.selectSheet('empty');
        assert(!(await r.read()), `xlspy->TS ${format}: empty sheet has no rows`);
    } finally {
        await r.close();
    }
    console.log(`xlspy write -> TS read ok: ${format}`);
}

// ---------------------------------------------------------------------------
// D. Updater round-trips (both directions, both formats)
// ---------------------------------------------------------------------------

const NEW_HEADER = ['id', 'when', 'note'];
const NEW_BRIDGE: BridgeRow[] = [
    [1, { datetime: '2025-06-30T12:00:00' }, 'x'],
    [2, { datetime: '1900-02-28' }, 'y'],
    [3, null, 'z'],
];
const NEW_WITH_HEADER: BridgeRow[] = [NEW_HEADER, ...NEW_BRIDGE];

async function testTsUpdateXlspyRead(format: 'xlsx' | 'xlsb'): Promise<void> {
    const src = path.join(outputDir, `upd-src.${format}`);
    bridge(['write', '--format', format, '--file', src], {
        sheets: [{ name: 'Data', rows: MATRIX_WITH_HEADER }],
    });
    const out = path.join(outputDir, `upd-ts-out.${format}`);
    if (format === 'xlsx') {
        const u = new XlsxUpdater(src);
        u.replaceSheetData('Data', toTsRows(NEW_BRIDGE), { headers: [...NEW_HEADER] });
        u.save(out);
    } else {
        const u = new XlsbUpdater(src);
        u.replaceSheetData('Data', toTsRows(NEW_BRIDGE), { headers: [...NEW_HEADER] });
        u.save(out);
    }
    const res = bridge(['read', '--file', out], {});
    assertRowsEqual((res.sheets as any)['Data'], NEW_WITH_HEADER, `xlspy-write TS-update ${format}`);
    console.log(`xlspy write -> TS update -> xlspy read ok: ${format}`);
}

async function testXlspyUpdateTsRead(format: 'xlsx' | 'xlsb'): Promise<void> {
    const src = path.join(outputDir, `upd-ts-src.${format}`);
    const W = format === 'xlsx' ? XlsxWriter : XlsbWriter;
    const w = new W(src);
    w.addSheet('Data');
    w.writeSheet(toTsRows(MATRIX), [...HEADER], false);
    await w.finalize();

    const out = path.join(outputDir, `upd-py-out.${format}`);
    const res = bridge(['replace-sheet-data', '--format', format, '--file', src, '--sheet', 'Data', '--out', out], {
        rows: NEW_BRIDGE,
        headers: NEW_HEADER,
    });
    assertRowsEqual((res.sheets as any)['Data'], NEW_WITH_HEADER, `xlspy update bridge re-read ${format}`);

    const R = format === 'xlsx' ? XlsxReader : XlsbReader;
    const sheets = await readAllSheets(new R(), out);
    assertRowsEqual(sheets['Data'], NEW_WITH_HEADER, `TS-write xlspy-update TS-read ${format}`);
    console.log(`TS write -> xlspy update -> TS read ok: ${format}`);
}

// ---------------------------------------------------------------------------
// E. Table growth across implementations (XLSX only; XLSB tables unpatched)
// ---------------------------------------------------------------------------

async function testTableGrowth(): Promise<void> {
    const src = path.join(outputDir, 'table-src.xlsx');
    bridge(['write', '--format', 'xlsx', '--file', src], {
        sheets: [{
            name: 'T',
            rows: [['A', 'B'], [1, 2], [3, 4]],
            tables: [{ displayName: 'Orders', ref: 'A1:B3', columns: ['A', 'B'] }],
        }],
    });
    const out = path.join(outputDir, 'table-ts-out.xlsx');
    const u = new XlsxUpdater(src);
    u.replaceSheetData('T', [[5, 6], [7, 8], [9, 0]], { headers: ['A', 'B'] });
    u.save(out);

    const zip = new AdmZip(out);
    const tableXml = zip.readAsText('xl/tables/table1.xml');
    assert(tableXml.includes('ref="A1:B4"'), `table ref grown to A1:B4, got ${tableXml.slice(0, 200)}`);

    const res = bridge(['read', '--file', out], {});
    assertRowsEqual((res.sheets as any)['T'], [['A', 'B'], [5, 6], [7, 8], [9, 0]], 'table-grown values');
    console.log('table growth ok: xlspy table -> TS update -> xlspy read');
}

// ---------------------------------------------------------------------------

async function main(): Promise<void> {
    if (!xlspyAvailable()) {
        console.log('SKIPPED: xlspy not available (needs Python with the xlspy package)');
        return;
    }
    const dataDir = xlspyRepoDir();
    if (dataDir) {
        await testGoldenFixtures(dataDir);
    } else {
        console.log('SKIPPED golden fixtures: xlspy repo checkout not found (set XLSPY_REPO)');
    }
    for (const fmt of ['xlsx', 'xlsb'] as const) {
        await testTsWriteXlspyRead(fmt);
        await testTsStreamingWriteXlspyRead(fmt);
        await testXlspyWriteTsRead(fmt);
        await testTsUpdateXlspyRead(fmt);
        await testXlspyUpdateTsRead(fmt);
    }
    await testTableGrowth();
    console.log('xlspy cross-validation tests passed');
}

main().catch((e) => { console.error(e); process.exitCode = 1; });
