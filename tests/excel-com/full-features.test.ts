/**
 * Broad Excel COM verification of everything browser/demo.html and
 * browser/test.html can generate.
 *
 * The two pages are executed in a child Node process, every download is
 * written to a temp directory, and Microsoft Excel opens each file and checks
 * the feature set: sheet names, tables, totals formulas, conditional-format
 * priorities, comments, protection flags and data validations.
 *
 * Run: $env:EXCEL_COM=1; npm run build:browser; npx ts-node tests/excel-com/full-features.test.ts
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { describe, it } from 'node:test';
import * as assert from 'node:assert/strict';

if (process.env.EXCEL_COM !== '1') {
    console.log('SKIP - set EXCEL_COM=1 to run Excel COM checks');
    process.exit(0);
}

const ROOT = path.resolve(__dirname, '..', '..');
const COM_DIR = __dirname;
const BUNDLE = path.join(ROOT, 'browser', 'browser-spreadsheet.js');
const PAGE_RUNNER = path.join(COM_DIR, 'run-browser-page.mjs');
const VERIFIER = path.join(COM_DIR, 'Verify-Generated.ps1');

interface ValidationExpectation {
    cell: string;
    type?: number;
    operator?: number;
    formula1?: string;
}

interface TableExpectation {
    name: string;
    showTotals?: boolean;
    totals?: Array<{ column: number; formula: string }>;
}

interface StyleExpectation {
    cell: string;
    fontName?: string;
    gradientStops?: number;
    pattern?: number;
    locked?: boolean;
    formulaHidden?: boolean;
    indent?: number;
}

interface SheetExpectation {
    name: string;
    tables?: number;
    cf?: number;
    cfPriorities?: boolean;
    cfFill?: number;
    comments?: number;
    protected?: boolean;
    allowInsertRows?: boolean;
    autofilter?: boolean;
    fontName?: string;
    usedRowsMin?: number;
    validations?: ValidationExpectation[];
    tableChecks?: TableExpectation[];
    styleChecks?: StyleExpectation[];
}

interface FileExpectation {
    label: string;
    path: string;
    sheetCount?: number;
    allSheetsAutofilter?: boolean;
    sheets?: SheetExpectation[];
}

interface VerifyResult {
    file: string;
    sheet: string;
    check: string;
    pass: boolean;
    expected: string | null;
    actual: string | null;
    message: string;
}

function runPage(mode: 'demo' | 'test', html: string, outDir: string): void {
    const run = spawnSync(process.execPath, [PAGE_RUNNER, mode, html, outDir, BUNDLE], {
        cwd: ROOT,
        encoding: 'utf-8',
        timeout: 180_000,
    });
    if (run.status !== 0) {
        throw new Error(`browser page ${mode} failed (${run.status}):\n${run.stdout}\n${run.stderr}`);
    }
}

function requireFile(dir: string, name: string): string {
    const file = path.join(dir, name);
    assert.ok(fs.existsSync(file), `${name} was generated`);
    assert.ok(fs.statSync(file).size > 0, `${name} is not empty`);
    return file;
}

function workbook(dir: string, name: string, sheets: SheetExpectation[]): FileExpectation {
    return { label: name, path: requireFile(dir, name), sheets };
}

function showcase(dir: string, format: 'xlsx' | 'xlsb'): FileExpectation {
    return workbook(dir, `showcase.${format}`, [
        {
            name: 'Dashboard',
            tables: 1,
            cf: 7,
            cfPriorities: true,
            cfFill: 13551615,
            comments: 2,
            protected: true,
            allowInsertRows: false,
            fontName: 'Calibri',
            validations: [{ cell: 'F2', type: 3, formula1: 'High;Medium;Low' }],
        },
        {
            name: 'Layout & Outline',
            tables: 0,
            cf: 0,
            comments: 0,
            protected: false,
            usedRowsMin: 6,
        },
        {
            name: 'Formats & Formulas',
            tables: 0,
            cf: 0,
            autofilter: true,
            usedRowsMin: 3,
        },
        {
            name: 'Conditional Formats',
            tables: 0,
            cf: 12,
            cfPriorities: true,
            autofilter: true,
            usedRowsMin: 6,
        },
        {
            name: 'Tables & Validation',
            tables: 1,
            cf: 0,
            comments: 1,
            protected: true,
            allowInsertRows: false,
            validations: [
                { cell: 'E2', type: 3, formula1: 'Open;Closed' },
                { cell: 'A9', type: 1, operator: 1 },
                { cell: 'E9', type: 6, operator: 8 },
                { cell: 'F9', type: 7 },
            ],
            tableChecks: [{
                name: 'RegionalTotals',
                showTotals: true,
                totals: [
                    { column: 2, formula: '=SUBTOTAL(109,B2:B5)' },
                    { column: 4, formula: '=SUM(D2:D5)' },
                ],
            }],
        },
    ]);
}

function features(dir: string, format: 'xlsx' | 'xlsb'): FileExpectation {
    return workbook(dir, `features.${format}`, [
        {
            name: 'Features',
            tables: 1,
            cf: 5,
            cfPriorities: true,
            cfFill: 13551615,
            comments: 2,
            protected: true,
            fontName: 'Calibri',
            validations: [{ cell: 'E2', type: 3, formula1: 'High;Low' }],
            tableChecks: [{ name: 'FeatureTable', showTotals: false }],
        },
        {
            name: 'Conditional Formats',
            tables: 0,
            cf: 12,
            cfPriorities: true,
        },
        {
            name: 'Tables & Validation',
            tables: 1,
            comments: 1,
            protected: true,
            tableChecks: [{ name: 'RegionalTotals', showTotals: true }],
        },
    ]);
}

function dataFormats(dir: string, format: 'xlsx' | 'xlsb'): FileExpectation {
    const names = [
        'Basic Types', 'Finances', 'Dates and Time', 'Scientific', 'Geographic',
        'Warehouse', 'Grades', 'Various Types', 'Employees', 'Sales',
        'Template', 'Notes', 'Percentages', 'Flags', 'Table of Contents',
    ];
    return workbook(dir, `data_formats.${format}`, names.map(name => ({ name })));
}

function testSuite(dir: string): FileExpectation[] {
    return [
        workbook(dir, 'test_edge_cases.xlsx', [{ name: 'EdgeCases' }]),
        workbook(dir, 'test_edge_cases.xlsb', [{ name: 'EdgeCases' }]),
        workbook(dir, 'formats_test.xlsb', [{ name: 'Formats' }]),
        workbook(dir, 'formats_test.xlsx', [{ name: 'Formats' }]),
        workbook(dir, 'multi_15_filter.xlsb',
            Array.from({ length: 15 }, (_, i) => ({ name: `Sheet${i + 1}`, autofilter: true }))),
        workbook(dir, 'multi_15_filter.xlsx',
            Array.from({ length: 15 }, (_, i) => ({ name: `Sheet${i + 1}`, autofilter: true }))),
        features(dir, 'xlsx'),
        features(dir, 'xlsb'),
        workbook(dir, 'ultimate_styles.xlsx', [{ name: 'Ultimate' }]),
        workbook(dir, 'ultimate_styles.xlsb', [
            {
                name: 'Ultimate Styles',
                usedRowsMin: 6,
                fontName: 'Calibri',
                styleChecks: [
                    { cell: 'C3', fontName: 'Georgia' },
                    { cell: 'A3', gradientStops: 2 },
                    { cell: 'F3', locked: false, formulaHidden: true },
                    { cell: 'B5', indent: 3 },
                    { cell: 'F5', pattern: 9 },
                ],
            },
        ]),
    ];
}

function demoSuite(dir: string): FileExpectation[] {
    return [
        workbook(dir, 'simple_report.xlsx', [{ name: 'Sheet1', usedRowsMin: 6 }]),
        workbook(dir, 'simple_report.xlsb', [{ name: 'Sheet1', usedRowsMin: 6 }]),
        workbook(dir, 'large_report.xlsx', [{ name: 'Sheet1', usedRowsMin: 10_001 }]),
        workbook(dir, 'large_report.xlsb', [{ name: 'Sheet1', usedRowsMin: 10_001 }]),
        workbook(dir, 'multi_sheet.xlsx', [
            { name: 'Sales', usedRowsMin: 3 },
            { name: 'Inventory', usedRowsMin: 3 },
        ]),
        workbook(dir, 'multi_sheet.xlsb', [
            { name: 'Sales', usedRowsMin: 3 },
            { name: 'Inventory', usedRowsMin: 3 },
        ]),
        workbook(dir, 'formats_demo.xlsx', [{ name: 'Sheet1', usedRowsMin: 24 }]),
        workbook(dir, 'formats_demo.xlsb', [{ name: 'Sheet1', usedRowsMin: 24 }]),
        dataFormats(dir, 'xlsx'),
        dataFormats(dir, 'xlsb'),
        showcase(dir, 'xlsx'),
        showcase(dir, 'xlsb'),
        workbook(dir, 'ultimate-styles.xlsx', [{ name: 'Ultimate Styles', usedRowsMin: 2 }]),
        workbook(dir, 'ultimate-styles.xlsb', [
            {
                name: 'Ultimate Styles',
                usedRowsMin: 6,
                fontName: 'Calibri',
                styleChecks: [
                    { cell: 'C3', fontName: 'Georgia' },
                    { cell: 'A3', gradientStops: 2 },
                    { cell: 'F3', locked: false, formulaHidden: true },
                    { cell: 'B5', indent: 3 },
                    { cell: 'F5', pattern: 9 },
                ],
            },
        ]),
    ];
}

describe('Excel COM: generated browser downloads', () => {
    it('opens every demo and test-page download and keeps its features', () => {
        assert.ok(fs.existsSync(BUNDLE),
            'browser/browser-spreadsheet.js exists; run npm run build:browser first');

        const outputDir = path.join(os.tmpdir(), `justybase-com-${process.pid}`);
        const demoDir = path.join(outputDir, 'demo');
        const testDir = path.join(outputDir, 'test');
        fs.rmSync(outputDir, { recursive: true, force: true });

        runPage('demo', path.join(ROOT, 'browser', 'demo.html'), demoDir);
        runPage('test', path.join(ROOT, 'browser', 'test.html'), testDir);

        const manifest = {
            files: [
                ...demoSuite(demoDir),
                ...testSuite(testDir),
            ],
        };
        const manifestPath = path.join(outputDir, 'manifest.json');
        const resultPath = path.join(outputDir, 'results.json');
        fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));

        const run = spawnSync('pwsh', [
            '-NoProfile', '-File', VERIFIER, manifestPath, resultPath,
        ], { cwd: ROOT, encoding: 'utf-8', timeout: 600_000 });
        const output = `${run.stdout}\n${run.stderr}`;
        assert.ok(fs.existsSync(resultPath), `verifier wrote results\n${output}`);

        const results = JSON.parse(fs.readFileSync(resultPath, 'utf-8')) as VerifyResult[];
        const failures = results.filter(result => !result.pass);
        const expectedFiles = new Set(manifest.files.map(file => file.label));
        const openedFiles = new Set(
            results.filter(result => result.check === 'open' && result.pass).map(result => result.file),
        );
        assert.deepEqual(
            [...expectedFiles].filter(file => !openedFiles.has(file)),
            [],
            'every generated file was opened by Excel',
        );
        assert.deepEqual(
            failures.map(result =>
                `${result.file}${result.sheet ? '/' + result.sheet : ''}: `
                + `${result.check} expected ${result.expected} got ${result.actual}`
                + (result.message ? ` (${result.message})` : '')),
            [],
            `Excel found ${failures.length} failing checks`,
        );
        assert.ok(results.length >= 100, `at least 100 feature checks ran (ran ${results.length})`);

        fs.rmSync(outputDir, { recursive: true, force: true });
    });
});
