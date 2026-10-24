/**
 * Excel COM conformance suite driven by the shared spec in comparison/com/cases.json.
 *
 * Writer tests: XlsxWriter / XlsbWriter write every case; Microsoft Excel (COM) opens each file and
 * checks structure (sheets, visibility, AutoFilter, used range) and values (every cell and number format).
 * Reader tests: Excel writes the same cases; XlsxReader / XlsbReader read them back.
 * Updater tests: an Excel-authored template is updated in place; Excel verifies the result.
 *
 * Run: npx ts-node tests/excel-com/conformance.test.ts  (Windows, Microsoft Excel, PowerShell 7 required)
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { describe, it, before } from 'node:test';
import * as assert from 'node:assert/strict';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsxUpdater } from '../../src/XlsxUpdater';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import type { CellValue } from '../../src/Formats';

const REPO_ROOT = path.resolve(__dirname, '..', '..', '..');
const COM_DIR = path.join(REPO_ROOT, 'comparison', 'com');
const SPEC_PATH = path.join(COM_DIR, 'cases.json');
const FORMATS = ['xlsx', 'xlsb'] as const;
type Format = (typeof FORMATS)[number];

// ─────────────────────────────────────────────────────────────────────────────
// Spec handling (mirror of Expand-SheetRows in comparison/com/ExcelComSpec.psm1)
// ─────────────────────────────────────────────────────────────────────────────

type SpecCell = null | string | number | boolean | { date: string } | { value: string | number | boolean; numberFormat?: string };

interface SpecSheet {
    name: string;
    hidden?: boolean;
    autoFilter?: boolean;
    rows?: SpecCell[][];
    generator?: { rows: number; header?: string[]; columns?: number };
}

interface SpecCase {
    id: string;
    sheets: SpecSheet[];
}

const spec: { cases: SpecCase[] } = JSON.parse(fs.readFileSync(SPEC_PATH, 'utf-8'));

function expandSheet(sheet: SpecSheet): SpecCell[][] {
    const gen = sheet.generator;
    if (!gen) return sheet.rows ?? [];
    if (gen.columns !== undefined) {
        const cols = gen.columns;
        const grid: SpecCell[][] = [Array.from({ length: cols }, (_, j) => `C${String(j + 1).padStart(2, '0')}`)];
        for (let i = 1; i <= gen.rows; i++) {
            grid.push(Array.from({ length: cols }, (_, j) => i * 1000 + (j + 1)));
        }
        return grid;
    }
    const grid: SpecCell[][] = [[...(gen.header ?? [])]];
    for (let i = 1; i <= gen.rows; i++) grid.push([i, `row-${i}`, i * 0.25]);
    return grid;
}

/** Spec date -> JS Date whose UTC components are the wall-clock time written in the spec (the TS writer serialises UTC). */
function specDateToDate(iso: string): Date {
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/.exec(iso);
    if (!m) throw new Error(`bad spec date ${iso}`);
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]));
}

function specCellToWriterValue(cell: SpecCell): CellValue {
    if (cell !== null && typeof cell === 'object') {
        if ('date' in cell) return specDateToDate(cell.date);
        if (cell.numberFormat !== undefined) return { value: cell.value as number, format: cell.numberFormat };
        return cell.value as CellValue;
    }
    return cell;
}

/** Expected value as the reader should return it, with dates as UTC-wall-clock Date objects. */
function specCellExpected(cell: SpecCell): unknown {
    if (cell !== null && typeof cell === 'object') {
        if ('date' in cell) return specDateToDate(cell.date);
        return cell.value;
    }
    return cell;
}

function excelSerialOfWallClock(d: Date): number {
    return (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function sameCell(expected: unknown, actual: unknown): boolean {
    if (expected === null || expected === undefined) return actual === null || actual === undefined || actual === '';
    if (typeof expected === 'boolean') return actual === expected;
    if (typeof expected === 'string') return actual === expected;
    if (typeof expected === 'number') {
        if (typeof actual === 'bigint') return Number(actual) === expected;
        return typeof actual === 'number' && Math.abs(actual - expected) <= 1e-9 * Math.max(1, Math.abs(expected));
    }
    if (expected instanceof Date) {
        if (actual instanceof Date) return Math.abs(actual.getTime() - expected.getTime()) <= 1000;
        // Raw serial numbers are accepted only when they hold the same wall-clock instant.
        if (typeof actual === 'number') return Math.abs(actual - excelSerialOfWallClock(expected)) <= 1 / 86400;
    }
    return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// Excel COM helpers (PowerShell scripts shared with the other libraries)
// ─────────────────────────────────────────────────────────────────────────────

interface VerifyTarget {
    file: string;
    caseId: string;
    format: Format;
}

interface VerifyResult {
    caseId: string;
    format: Format;
    file: string;
    sheet: string;
    check: string;
    pass: boolean;
    expected: string | null;
    actual: string | null;
    message: string;
}

function runPwsh(script: string, args: string[]): { status: number | null; stdout: string; stderr: string } {
    const result = spawnSync(
        'pwsh',
        ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(COM_DIR, script), ...args],
        { encoding: 'utf-8', timeout: 900_000, maxBuffer: 64 * 1024 * 1024 }
    );
    return { status: result.status, stdout: result.stdout ?? '', stderr: result.stderr ?? '' };
}

function verifyWithExcel(targets: VerifyTarget[], outDir: string): VerifyResult[] {
    const targetsPath = path.join(outDir, `verify-targets-${Date.now()}.json`);
    const resultsPath = path.join(outDir, `verify-results-${Date.now()}.json`);
    fs.writeFileSync(targetsPath, JSON.stringify(targets));
    const run = runPwsh('Invoke-ExcelVerify.ps1', ['-SpecPath', SPEC_PATH, '-TargetsPath', targetsPath, '-OutPath', resultsPath]);
    if (!fs.existsSync(resultsPath)) {
        throw new Error(`Excel verifier produced no results.\n${run.stdout}\n${run.stderr}`);
    }
    return JSON.parse(fs.readFileSync(resultsPath, 'utf-8')) as VerifyResult[];
}

function excelAvailable(): boolean {
    if (process.platform !== 'win32') return false;
    const reg = spawnSync('reg', ['query', 'HKCR\\Excel.Application\\CLSID'], { encoding: 'utf-8' });
    const pwsh = spawnSync('pwsh', ['-NoProfile', '-Command', '$PSVersionTable.PSVersion.Major'], { encoding: 'utf-8' });
    return reg.status === 0 && pwsh.status === 0;
}

function failureText(results: VerifyResult[]): string {
    return results
        .filter(r => !r.pass)
        .map(r => `${r.check} [${r.sheet}]: expected '${r.expected}', got '${r.actual}' ${r.message}`)
        .join('\n');
}

const isStructureCheck = (check: string) => !check.startsWith('cells:') && !check.startsWith('numberFormat:');

// ─────────────────────────────────────────────────────────────────────────────
// Session: writes every case, verifies with Excel, and creates Excel fixtures once.
// ─────────────────────────────────────────────────────────────────────────────

const available = excelAvailable();
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'justybase-com-'));
const writtenDir = path.join(workDir, 'written');
const fixtureDir = path.join(workDir, 'excel-fixtures');
let writerResults: VerifyResult[] = [];
let fixtures = new Map<string, string>();

async function writeCase(c: SpecCase, file: string, format: Format): Promise<void> {
    const writer = format === 'xlsx' ? new XlsxWriter(file) : new XlsbWriter(file);
    for (const sheet of c.sheets) {
        writer.addSheet(sheet.name, sheet.hidden === true);
        const rows = expandSheet(sheet).map(row => row.map(specCellToWriterValue));
        // writeSheet adds an AutoFilter only when headers are passed, so AutoFilter sheets use row 1 as headers.
        if (sheet.autoFilter === true && rows.length > 0) {
            writer.writeSheet(rows.slice(1), rows[0] as string[], true);
        } else {
            writer.writeSheet(rows, null, false);
        }
    }
    await writer.finalize();
}

function caseFormatPairs(): Array<[string, Format]> {
    return spec.cases.flatMap(c => FORMATS.map(f => [c.id, f] as [string, Format]));
}

before(async () => {
    if (!available) return;
    fs.mkdirSync(writtenDir, { recursive: true });
    const targets: VerifyTarget[] = [];
    for (const c of spec.cases) {
        for (const format of FORMATS) {
            const file = path.join(writtenDir, `${c.id}.${format}`);
            await writeCase(c, file, format);
            targets.push({ file, caseId: c.id, format });
        }
    }
    writerResults = verifyWithExcel(targets, workDir);

    const fixtureRun = runPwsh('New-ExcelFixture.ps1', ['-SpecPath', SPEC_PATH, '-OutDir', fixtureDir]);
    if (fixtureRun.status !== 0) throw new Error(`New-ExcelFixture failed:\n${fixtureRun.stdout}\n${fixtureRun.stderr}`);
    const fixtureTargets: VerifyTarget[] = JSON.parse(fs.readFileSync(path.join(fixtureDir, 'targets.json'), 'utf-8'));
    fixtures = new Map(fixtureTargets.map(t => [`${t.caseId}|${t.format}`, t.file]));
});

// ─────────────────────────────────────────────────────────────────────────────
// Tests
// ─────────────────────────────────────────────────────────────────────────────

function checksFor(caseId: string, format: Format, include: (check: string) => boolean): VerifyResult[] {
    return writerResults.filter(r => r.caseId === caseId && r.format === format && include(r.check));
}

async function readSheetRows(reader: XlsxReader | XlsbReader, sheetName: string): Promise<unknown[][]> {
    await reader.selectSheet(sheetName);
    const rows: unknown[][] = [];
    while (await reader.read()) {
        const row: unknown[] = [];
        const width = Math.max(reader.fieldCount, 0);
        for (let i = 0; i < width; i++) row.push(reader.getValue(i));
        rows.push(row);
    }
    return rows;
}

const suite = available ? describe : describe.skip;

suite('Excel COM conformance', () => {
    for (const [caseId, format] of caseFormatPairs()) {
        describe(`${caseId} (${format})`, () => {
            it('writer: structure matches what Excel sees', () => {
                const results = checksFor(caseId, format, isStructureCheck);
                assert.ok(results.length > 0, 'Excel produced no structure checks for this file');
                assert.equal(failureText(results), '');
            });

            it('writer: every cell value and number format matches', () => {
                const results = checksFor(caseId, format, c => !isStructureCheck(c));
                assert.ok(results.length > 0, 'Excel produced no cell checks for this file');
                assert.equal(failureText(results), '');
            });

            it('reader: reads the Excel-authored file as specified', async () => {
                const file = fixtures.get(`${caseId}|${format}`);
                assert.ok(file, 'Excel fixture was not created');
                const c = spec.cases.find(x => x.id === caseId)!;
                const reader = format === 'xlsx' ? new XlsxReader() : new XlsbReader();
                await reader.open(file);
                try {
                    assert.deepEqual(reader.getSheetNames(), c.sheets.map(s => s.name));
                    for (const sheet of c.sheets) {
                        const expected = expandSheet(sheet).map(row => row.map(specCellExpected));
                        if (expected.length === 0) continue;
                        const width = Math.max(...expected.map(r => r.length));
                        const actual = await readSheetRows(reader, sheet.name);
                        assert.ok(actual.length >= expected.length, `${sheet.name}: expected at least ${expected.length} rows, got ${actual.length}`);
                        for (let r = 0; r < expected.length; r++) {
                            for (let col = 0; col < width; col++) {
                                const exp = col < expected[r].length ? expected[r][col] : null;
                                const act = actual[r] && col < actual[r].length ? actual[r][col] : null;
                                assert.ok(sameCell(exp, act), `${sheet.name} row ${r + 1} col ${col + 1}: expected ${String(exp)}, got ${String(act)}`);
                            }
                        }
                    }
                } finally {
                    await reader.close();
                }
            });
        });
    }

    describe('updater on an Excel-authored template', () => {
        for (const format of FORMATS) {
            it(`${format}: replaces sheet Data, keeps sheet Keep, Excel verifies`, async () => {
                const template = fixtures.get(`update_template|${format}`);
                assert.ok(template, 'template fixture was not created');
                const target = path.join(workDir, `updated.${format}`);
                fs.copyFileSync(template, target);

                const resultCase = spec.cases.find(x => x.id === 'update_result')!;
                const rows = expandSheet(resultCase.sheets[0]).map(row => row.map(specCellToWriterValue)) as CellValue[][];
                const headers = (rows[0] as string[]);
                const body = rows.slice(1);

                if (format === 'xlsx') {
                    const updater = new XlsxUpdater(target);
                    try {
                        updater.replaceSheetData('Data', body, { headers });
                        updater.save(target);
                    } finally {
                        updater.dispose();
                    }
                } else {
                    const updater = new XlsbUpdater(target);
                    try {
                        updater.replaceSheetData('Data', body, { headers });
                        updater.save(target);
                    } finally {
                        updater.dispose();
                    }
                }

                const results = verifyWithExcel([{ file: target, caseId: 'update_result', format }], workDir);
                assert.equal(failureText(results), '');
            });
        }
    });
});
