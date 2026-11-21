/**
 * Regression tests for the Excel-compatibility fixes:
 *   - differential fills use bgColor and borders omit empty start/end elements
 *   - sheet-protection flags block unspecified actions in both formats
 *   - conditional-format priorities restart on every sheet
 *   - XLSB date styles map date/date-time to the right base XF
 *   - XLSB icon sets store one CFVO per icon (0% baseline first)
 *   - XLSB data-validation operators use the BIFF12 order
 *   - table names cannot be cell references; built-in totals get SUBTOTAL
 *   - the legacy password hash matches Excel's algorithm
 *   - CF/DVal formula references are origin-relative (PtgRefN/PtgAreaN)
 *   - XLSB packages carry the Office theme (Calibri instead of Aptos)
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as assert from 'node:assert/strict';
import AdmZip from 'adm-zip';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import {
    normalizeTableSpec,
    passwordHash,
    tableTotalsOverrides,
} from '../../src/writers/core/sheetLayout';
import { buildXlsbSheetProtectionRecord } from '../../src/writers/core/xlsbLayout';
import {
    XlsbDxfRegistry,
    buildXlsbConditionalFormatRecords,
} from '../../src/writers/core/xlsbConditionalFormat';
import { XlsbStyleRegistry } from '../../src/writers/core/xlsbStyles';
import { buildXlsbValidationRecords } from '../../src/writers/core/xlsbSheetFeatures';
import { compileXlsbFormulaStats } from '../../src/writers/core/xlsbFormula';
import { CellStyle } from '../../src/Formats';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-regressions-'));

interface Biff12Record {
    id: number;
    body: Buffer;
}

/** Parse a BIFF12 record stream (used to inspect raw XLSB records). */
function scanRecords(data: Buffer): Biff12Record[] {
    const records: Biff12Record[] = [];
    let position = 0;
    while (position < data.length) {
        let id = 0;
        let shift = 0;
        let byte: number;
        do {
            byte = data[position++];
            id |= (byte & 0x7f) << shift;
            shift += 7;
        } while (byte & 0x80);
        let length = 0;
        shift = 0;
        do {
            byte = data[position++];
            length |= (byte & 0x7f) << shift;
            shift += 7;
        } while (byte & 0x80);
        records.push({ id, body: data.subarray(position, position + length) });
        position += length;
    }
    return records;
}

function unpack(file: string): Map<string, Buffer> {
    const parts = new Map<string, Buffer>();
    for (const entry of new AdmZip(file).getEntries()) {
        if (!entry.isDirectory) parts.set(entry.entryName, entry.getData());
    }
    return parts;
}

function partText(parts: Map<string, Buffer>, name: string): string {
    const part = parts.get(name);
    assert.ok(part, `part ${name} exists`);
    return part!.toString('utf8');
}

const BASIC_STYLE: CellStyle = {
    border: {
        left: { style: 'thin', color: 'FF000000' },
        right: { style: 'thin', color: 'FF000000' },
        top: { style: 'thin', color: 'FF000000' },
        bottom: { style: 'thin', color: 'FF000000' },
    },
};

const CONDITIONAL_FORMAT = (sqref: string, count = 1) => ({
    sqref,
    rules: Array.from({ length: count }, (_, index) => ({
        type: 'cellIs' as const,
        operator: 'greaterThan' as const,
        formula: [String(index)],
        style: { font: { bold: true } },
    })),
});

async function testXlsxDifferentialFillAndBorders(): Promise<void> {
    const file = path.join(outputDir, 'dxf.xlsx');
    const writer = new XlsxWriter(file, { compressionLevel: 0 });
    writer.addSheet('S');
    writer.writeSheet([[{ value: 1, style: BASIC_STYLE }]], ['H'], {
        conditionalFormats: [{
            sqref: 'A2:A2',
            rules: [{
                type: 'cellIs',
                operator: 'greaterThan',
                formula: ['0'],
                style: {
                    font: { bold: true },
                    fill: { patternType: 'solid', fgColor: 'FFFFC7CE' },
                },
            }],
        }],
    });
    await writer.finalize();

    const parts = unpack(file);
    const styles = partText(parts, 'xl/styles.xml');
    const dxf = /<dxfs[\s\S]*?<\/dxfs>/.exec(styles)?.[0] ?? '';
    assert.ok(dxf.includes('<bgColor rgb="FFFFC7CE"/>'), 'dxf solid fill uses bgColor');
    assert.ok(!dxf.includes('<fgColor rgb="FFFFC7CE"/>'), 'dxf solid fill does not use fgColor');

    const border = /<borders[\s\S]*?<\/borders>/.exec(styles)?.[0] ?? '';
    assert.ok(border.includes('<left style="thin">'), 'left border is written');
    assert.ok(border.includes('<right style="thin">'), 'right border is written');
    assert.ok(!border.includes('<start'), 'empty start element is omitted');
    assert.ok(!border.includes('<end'), 'empty end element is omitted');
    assert.ok(!border.includes('<vertical'), 'empty vertical element is omitted');
    assert.ok(!border.includes('<horizontal'), 'empty horizontal element is omitted');
}

async function testProtectionFlags(): Promise<void> {
    // XLSX: true blocks, false explicitly allows, undefined keeps the default.
    const files = {
        blocked: path.join(outputDir, 'protect-blocked.xlsx'),
        allowed: path.join(outputDir, 'protect-allowed.xlsx'),
        defaulted: path.join(outputDir, 'protect-defaulted.xlsx'),
    };
    const specs = [
        { file: files.blocked, insertRows: true },
        { file: files.allowed, insertRows: false },
        { file: files.defaulted, insertRows: undefined },
    ];
    for (const spec of specs) {
        const writer = new XlsxWriter(spec.file, { compressionLevel: 0 });
        writer.addSheet('S');
        writer.writeSheet([[1]], ['H'], {
            sheetProtection: { enabled: true, insertRows: spec.insertRows },
        });
        await writer.finalize();
    }
    const blocked = partText(unpack(files.blocked), 'xl/worksheets/sheet1.xml');
    const allowed = partText(unpack(files.allowed), 'xl/worksheets/sheet1.xml');
    const defaulted = partText(unpack(files.defaulted), 'xl/worksheets/sheet1.xml');
    assert.ok(blocked.includes('insertRows="1"'), 'XLSX insertRows: true blocks');
    assert.ok(allowed.includes('insertRows="0"'), 'XLSX insertRows: false allows');
    assert.ok(!defaulted.includes('insertRows='), 'XLSX unspecified insertRows stays default');

    const record = buildXlsbSheetProtectionRecord({
        enabled: true,
        formatCells: true,
        insertRows: true,
        selectLockedCells: true,
        selectUnlockedCells: false,
        objects: false,
    });
    const body = scanRecords(Buffer.from(record))[0].body;
    const flags: number[] = [];
    for (let index = 0; index < 16; index++) flags.push(body.readUInt32LE(2 + index * 4));
    assert.deepEqual(flags, [1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0],
        'XLSB blocks unspecified actions; only explicit false allows them');
}

async function testConditionalFormatPriorities(): Promise<void> {
    const xlsx = path.join(outputDir, 'priorities.xlsx');
    const xlsxWriter = new XlsxWriter(xlsx, { compressionLevel: 0 });
    xlsxWriter.addSheet('First');
    xlsxWriter.writeSheet([[1]], ['H'], {
        conditionalFormats: [CONDITIONAL_FORMAT('A2:A2', 2)],
    });
    xlsxWriter.addSheet('Second');
    xlsxWriter.writeSheet([[1]], ['H'], {
        conditionalFormats: [CONDITIONAL_FORMAT('A2:A2', 3)],
    });
    await xlsxWriter.finalize();
    const xlsxParts = unpack(xlsx);
    const firstXml = partText(xlsxParts, 'xl/worksheets/sheet1.xml');
    const secondXml = partText(xlsxParts, 'xl/worksheets/sheet2.xml');
    assert.deepEqual([...firstXml.matchAll(/priority="(\d+)"/g)].map(match => Number(match[1])), [1, 2],
        'XLSX first sheet priorities start at 1');
    assert.deepEqual([...secondXml.matchAll(/priority="(\d+)"/g)].map(match => Number(match[1])), [1, 2, 3],
        'XLSX second sheet priorities restart at 1');

    const xlsb = path.join(outputDir, 'priorities.xlsb');
    const xlsbWriter = new XlsbWriter(xlsb, { compressionLevel: 0 });
    xlsbWriter.addSheet('First');
    xlsbWriter.writeSheet([[1]], ['H'], {
        conditionalFormats: [CONDITIONAL_FORMAT('A2:A2', 2)],
    });
    xlsbWriter.addSheet('Second');
    xlsbWriter.writeSheet([[1]], ['H'], {
        conditionalFormats: [CONDITIONAL_FORMAT('A2:A2', 3)],
    });
    await xlsbWriter.finalize();
    const xlsbParts = unpack(xlsb);
    const priorities = (name: string): number[] => scanRecords(xlsbParts.get(name)!)
        .filter(record => record.id === 0x1cf)
        .map(record => record.body.readUInt32LE(12));
    assert.deepEqual(priorities('xl/worksheets/sheet1.bin'), [1, 2],
        'XLSB first sheet priorities start at 1');
    assert.deepEqual(priorities('xl/worksheets/sheet2.bin'), [1, 2, 3],
        'XLSB second sheet priorities restart at 1');
}

async function testXlsbDateStyles(): Promise<void> {
    const file = path.join(outputDir, 'dates.xlsb');
    const writer = new XlsbWriter(file, { compressionLevel: 0 });
    writer.addSheet('Dates');
    writer.writeSheet([
        [new Date(Date.UTC(2024, 0, 2))],
        [new Date(Date.UTC(2024, 0, 2, 15, 30))],
    ], ['Date']);
    await writer.finalize();

    const sheet = unpack(file).get('xl/worksheets/sheet1.bin')!;
    const styles = scanRecords(sheet)
        .filter(record => record.id === 0x05)
        .map(record => record.body.readUInt32LE(4) & 0xff);
    assert.deepEqual(styles, [2, 1],
        'date-only cells use the ISO date XF, date-time cells use the date-time XF');
}

function testIconSetBaseline(): void {
    const dxfs = new XlsbDxfRegistry(() => 164);
    for (const [iconSet, expected] of [
        ['3TrafficLights1', 3],
        ['4Arrows', 4],
        ['5Arrows', 5],
    ] as const) {
        const records = buildXlsbConditionalFormatRecords([{
            sqref: 'E2:E6',
            rules: [{
                type: 'iconSet',
                iconSet,
                cfvo: [
                    { type: 'percent', value: '33' },
                    { type: 'percent', value: '67' },
                ],
            }],
        }] as any, dxfs);
        const thresholds = scanRecords(Buffer.from(records)).filter(record => record.id === 0x1d7);
        assert.equal(thresholds.length, expected, `${iconSet} stores one CFVO per icon`);
        assert.equal(thresholds[0].body.readUInt32LE(0), 4, `${iconSet} baseline is percent`);
        assert.equal(thresholds[0].body.readDoubleLE(4), 0, `${iconSet} baseline is 0%`);
    }
}

function testDataValidationOperators(): void {
    const operators = ['greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const;
    const records = buildXlsbValidationRecords(operators.map(operator => ({
        sqref: 'A1',
        type: 'whole',
        operator,
        formula1: '1',
    })));
    const codes = scanRecords(Buffer.from(records))
        .filter(record => record.id === 0x40)
        .map(record => (record.body.readUInt32LE(0) >>> 20) & 0x0f);
    assert.deepEqual(codes, [4, 5, 6, 7], 'BIFF12 data-validation operator codes');
}

function testTableNamesAndTotals(): void {
    for (const name of ['A1', 'T1', 'R1C1', 'R', 'C', 'XFD1048576']) {
        assert.throws(
            () => normalizeTableSpec({ name, ref: 'A1:B5', columns: ['A', 'B'] }, 1),
            /Invalid table name/,
            `${name} is rejected as a table name`,
        );
    }
    assert.ok(normalizeTableSpec({ name: 'Sales', ref: 'A1:B5', columns: ['A', 'B'] }, 1),
        'an ordinary table name is accepted');

    const table = normalizeTableSpec({
        name: 'Totals',
        ref: 'A1:C5',
        totalsRow: true,
        columns: [
            { name: 'Label', totalsRowLabel: 'Total' },
            { name: 'Q1', totalsRowFunction: 'sum' },
            { name: 'Double', totalsRowFunction: 'custom', totalsRowFormula: '=SUM(C2:C4)*2' },
        ],
    }, 1);
    assert.deepEqual(tableTotalsOverrides([table]), [
        { row: 4, col: 0, text: 'Total' },
        { row: 4, col: 1, formula: 'SUBTOTAL(109,B2:B4)' },
        { row: 4, col: 2, formula: 'SUM(C2:C4)*2' },
    ], 'built-in totals produce SUBTOTAL formulas, custom totals are kept');
}

function testPasswordHash(): void {
    assert.equal(passwordHash('password'), 0x83af, 'Excel hash for password');
    assert.equal(passwordHash('secret'), 0xdaa7, 'Excel hash for secret');
    assert.equal(passwordHash('demo'), 0xca8b, 'Excel hash for demo');
}

function testRelativeFormulaTokens(): void {
    const cell = compileXlsbFormulaStats('A2>0', { origin: { col: 0, row: 0 } });
    assert.equal(cell.tokens[0], 0x44, 'cell formulas use PtgRef');

    const relative = compileXlsbFormulaStats('A2>0', {
        origin: { col: 0, row: 0 },
        relativeOffsets: true,
    });
    assert.equal(relative.tokens[0], 0x4c, 'CF/DVal formulas use PtgRefN');
    assert.equal(Buffer.from(relative.tokens).readUInt32LE(1), 1, 'row offset is origin-relative');

    const area = compileXlsbFormulaStats('A2:B3', {
        origin: { col: 0, row: 1 },
        relativeOffsets: true,
    });
    assert.equal(area.tokens[0], 0x4d, 'CF/DVal ranges use PtgAreaN');
    assert.equal(Buffer.from(area.tokens).readUInt32LE(1), 0, 'area start row is origin-relative');
    assert.equal(Buffer.from(area.tokens).readUInt32LE(5), 1, 'area end row is origin-relative');
}

async function testThemePart(): Promise<void> {
    const file = path.join(outputDir, 'theme.xlsb');
    const writer = new XlsbWriter(file, { compressionLevel: 0 });
    writer.addSheet('S');
    writer.writeSheet([[1]], ['H']);
    await writer.finalize();

    const parts = unpack(file);
    const theme = partText(parts, 'xl/theme/theme1.xml');
    assert.ok(theme.startsWith('<?xml'), 'theme starts with the XML declaration');
    assert.ok(theme.includes('typeface="Calibri"'), 'theme declares Calibri');
    assert.ok(theme.trimEnd().endsWith('</a:theme>'), 'theme is complete');
    assert.ok(partText(parts, '[Content_Types].xml').includes('/xl/theme/theme1.xml'),
        'theme content type is registered');
    assert.ok(partText(parts, 'xl/_rels/workbook.bin.rels').includes('theme/theme1.xml'),
        'theme relationship is registered');
}

async function testXlsbRejectsUnsupportedStyles(): Promise<void> {
    const cases: Array<[string, CellStyle]> = [
        ['relativeIndent', { alignment: { relativeIndent: 2 } }],
        ['start border', { border: { start: { style: 'thin' } } }],
        ['end border', { border: { end: { style: 'thin' } } }],
        ['vertical border', { border: { vertical: { style: 'thin' } } }],
        ['horizontal border', { border: { horizontal: { style: 'thin' } } }],
        ['quotePrefix', { quotePrefix: true }],
        ['pivotButton', { pivotButton: true }],
    ];
    for (const [label, style] of cases) {
        assert.throws(
            () => new XlsbStyleRegistry().register(style),
            /no record slot/,
            `XLSB rejects ${label}`,
        );
    }
}

async function main(): Promise<void> {
    try {
        await testXlsxDifferentialFillAndBorders();
        await testProtectionFlags();
        await testConditionalFormatPriorities();
        await testXlsbDateStyles();
        testIconSetBaseline();
        testDataValidationOperators();
        testTableNamesAndTotals();
        testPasswordHash();
        testRelativeFormulaTokens();
        await testThemePart();
        await testXlsbRejectsUnsupportedStyles();
        console.log('Regression tests passed');
    } finally {
        fs.rmSync(outputDir, { recursive: true, force: true });
    }
}

main().catch(error => {
    console.error(error);
    process.exit(1);
});
