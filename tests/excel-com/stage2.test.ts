/**
 * Excel COM checks for the Etap 2 features (styles, formulas, merges,
 * hyperlinks, validations, conditional formatting, comments, widths).
 *
 * Run: $env:EXCEL_COM=1; npx ts-node tests/excel-com/stage2.test.ts
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { describe, it, before } from 'node:test';
import * as assert from 'node:assert/strict';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { CellStyle, CellValue, FormulaCell, RichText, StyledCell } from '../../src/Formats';

if (process.env.EXCEL_COM !== '1') {
    console.log('SKIP - set EXCEL_COM=1 to run Excel COM checks');
    process.exit(0);
}

function excelAvailable(): boolean {
    const reg = spawnSync('reg', ['query', 'HKCR\\Excel.Application\\CLSID'], { encoding: 'utf-8' });
    const pwsh = spawnSync('pwsh', ['-NoProfile', '-Command', '$PSVersionTable.PSVersion.Major'], { encoding: 'utf-8' });
    return reg.status === 0 && pwsh.status === 0;
}

interface ExcelResult {
    ok: boolean;
    error?: string;
    valueA2?: number;
    formatA2?: string;
    boldA2?: boolean;
    fontColorA2?: number;
    fillA2?: number;
    borderTopA2?: number;
    alignA2?: number;
    formulaB2?: string;
    valueB2?: number;
    valueB3?: number;
    textC2?: string;
    merged?: boolean;
    links?: number;
    validationType?: number;
    validationFormula?: string;
    conditions?: number;
    widthA?: number;
    widthB?: number;
    comment?: string | null;
}

function verify(file: string): ExcelResult {
    const script = path.join(__dirname, 'Verify-Stage2.ps1');
    const run = spawnSync('pwsh', ['-NoProfile', '-File', script, file], { encoding: 'utf-8' });
    if (run.status !== 0) throw new Error(`PowerShell verifier failed:\n${run.stdout}\n${run.stderr}`);
    const line = run.stdout.trim().split(/\r?\n/).filter(Boolean).pop();
    if (!line) throw new Error(`Empty verifier output:\n${run.stdout}\n${run.stderr}`);
    return JSON.parse(line) as ExcelResult;
}

const STYLE: CellStyle = {
    numFmt: '#,##0.00',
    font: { bold: true, color: 'FFFF0000' },
    fill: { patternType: 'solid', fgColor: 'FFFFFF00' },
    border: { top: { style: 'thin', color: 'FF000000' } },
    alignment: { horizontal: 'center' },
};

const ROWS: CellValue[][] = [
    [
        { value: 10, style: STYLE } as StyledCell,
        { formula: 'SUM(A2:A3)', value: 30 } as FormulaCell,
        new RichText([
            { text: 'red', style: { font: { bold: true, color: 'FFFF0000' } } },
            { text: ' plain' },
        ]),
    ],
    [20, { formula: 'A2*2', value: 20 } as FormulaCell, 'x'],
];

const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-stage2-com-'));

async function writeXlsx(): Promise<string> {
    const file = path.join(workDir, 'stage2.xlsx');
    const writer = new XlsxWriter(file);
    writer.addSheet('Feat');
    writer.writeSheet(ROWS, ['A', 'B', 'C'], {
        merges: ['A4:B4'],
        hyperlinks: [
            { ref: 'C1', target: 'https://example.com', tooltip: 'Example' },
            { ref: 'A2', location: 'Feat!A1' },
        ],
        dataValidations: [{ sqref: 'C2:C6', type: 'list', values: ['x', 'y'], allowBlank: true }],
        conditionalFormats: [{
            sqref: 'A2:A3',
            rules: [{
                type: 'cellIs',
                operator: 'greaterThan',
                formula: ['15'],
                style: { font: { bold: true }, fill: { fgColor: 'FFFFC7CE' } },
            }],
        }],
        comments: [{ ref: 'A2', text: 'note text', author: 'Tester' }],
        columnWidths: [30, 40],
    });
    await writer.finalize();
    return file;
}

async function writeXlsb(): Promise<string> {
    const file = path.join(workDir, 'stage2.xlsb');
    const plainRows: CellValue[][] = ROWS.map(row => row.map(cell =>
        cell instanceof RichText ? cell.toString() : cell));
    const writer = new XlsbWriter(file);
    writer.addSheet('Feat');
    writer.writeSheet(plainRows, ['A', 'B', 'C'], {
        merges: ['A4:B4'],
        hyperlinks: [
            { ref: 'C1', target: 'https://example.com', tooltip: 'Example' },
            { ref: 'A2', location: 'Feat!A1' },
        ],
        dataValidations: [{ sqref: 'C2:C6', type: 'list', values: ['x', 'y'], allowBlank: true }],
        columnWidths: [30, 40],
    });
    await writer.finalize();
    return file;
}

describe('Excel COM stage 2 features', () => {
    let available = false;
    before(() => {
        available = excelAvailable();
        if (!available) console.log('SKIP - Microsoft Excel / PowerShell 7 not available');
    });

    it('XLSX: styles, formulas, merges, links, validation, CF, comments, widths', async () => {
        if (!available) return;
        const result = verify(await writeXlsx());
        assert.equal(result.ok, true, result.error);
        assert.equal(result.valueA2, 10);
        assert.equal(result.formatA2, '#,##0.00');
        assert.equal(result.boldA2, true);
        assert.equal(result.fontColorA2, 255);
        assert.equal(result.fillA2, 65535);
        assert.equal(result.borderTopA2, 1);
        assert.equal(result.alignA2, -4108);
        assert.equal(result.formulaB2, '=SUM(A2:A3)');
        assert.equal(result.valueB2, 30);
        assert.equal(result.valueB3, 20);
        assert.equal(result.textC2, 'red plain');
        assert.equal(result.merged, true);
        assert.equal(result.links, 2);
        assert.equal(result.validationType, 3);
        assert.equal(result.validationFormula, 'x;y');
        assert.ok((result.conditions ?? 0) >= 1);
        assert.equal(result.comment, 'note text');
        assert.ok((result.widthA ?? 0) > 25 && (result.widthA ?? 0) <= 30.5, `width A ${result.widthA}`);
        assert.ok((result.widthB ?? 0) > 35 && (result.widthB ?? 0) <= 40.5, `width B ${result.widthB}`);
    });

    it('XLSB: styles, formulas, merges, links, validation, widths', async () => {
        if (!available) return;
        const result = verify(await writeXlsb());
        assert.equal(result.ok, true, result.error);
        assert.equal(result.valueA2, 10);
        assert.equal(result.formatA2, '#,##0.00');
        assert.equal(result.boldA2, true);
        assert.equal(result.fontColorA2, 255);
        assert.equal(result.fillA2, 65535);
        assert.equal(result.borderTopA2, 1);
        assert.equal(result.alignA2, -4108);
        assert.equal(result.formulaB2, '=SUM(A2:A3)');
        assert.equal(result.valueB2, 30);
        assert.equal(result.valueB3, 20);
        assert.equal(result.textC2, 'red plain');
        assert.equal(result.merged, true);
        assert.equal(result.links, 2);
        assert.equal(result.validationType, 3);
        assert.equal(result.validationFormula, 'x;y');
        assert.ok((result.widthA ?? 0) > 25 && (result.widthA ?? 0) <= 30.5, `width A ${result.widthA}`);
    });
});
