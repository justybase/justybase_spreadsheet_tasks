/**
 * Etap 2 feature tests: styles, formulas, rich text, sheet features and
 * column widths. Node and browser writers are compared byte-for-byte by part.
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as assert from 'node:assert/strict';
import AdmZip from 'adm-zip';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { BrowserXlsxWriter } from '../../src/browser/BrowserXlsxWriter';
import { BrowserXlsbWriter } from '../../src/browser/BrowserXlsbWriter';
import { CellStyle, CellValue, ExcelError, FormulaCell, RichText, StyledCell } from '../../src/Formats';
import {
    AUTOFILTER_BUTTON_WIDTH,
    AUTOFIT_SAMPLE_ROWS,
    COLUMN_WIDTH_PADDING,
    DEFAULT_COLUMN_WIDTH,
    MAX_AUTOFIT_WIDTH,
    autofitFileWidth,
    autofitWidth,
    textDisplayUnits,
    updateColWidthsFromRows,
} from '../../src/writerHelpers';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-stage2-'));

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
const HEADERS = ['A', 'B', 'C'];

/** XLSB does not emit rich strings yet; use the plain text equivalent. */
const XLSB_ROWS: CellValue[][] = ROWS.map(row =>
    row.map(cell => (cell instanceof RichText ? cell.toString() : cell)),
);

const FEATURES = {
    merges: ['A4:B4'],
    hyperlinks: [
        { ref: 'C1', target: 'https://example.com', tooltip: 'Example' },
        { ref: 'A2', location: 'Feat!A1' },
    ],
    dataValidations: [
        { sqref: 'C1:C5', type: 'list' as const, values: ['x', 'y'], allowBlank: true },
    ],
    conditionalFormats: [
        {
            sqref: 'A2:A3',
            rules: [{
                type: 'cellIs' as const,
                operator: 'greaterThan' as const,
                formula: ['15'],
                style: { font: { bold: true }, fill: { fgColor: 'FFFFC7CE' } },
            }],
        },
    ],
    comments: [{ ref: 'A1', text: 'note text', author: 'Tester' }],
    columnWidths: [30, 40],
};

function unpack(buffer: Buffer): Map<string, Buffer> {
    const parts = new Map<string, Buffer>();
    for (const entry of new AdmZip(buffer).getEntries()) {
        if (!entry.isDirectory) parts.set(entry.entryName, entry.getData());
    }
    return parts;
}

function partText(parts: Map<string, Buffer>, name: string): string {
    const part = parts.get(name);
    assert.ok(part, `part ${name} exists`);
    return part!.toString('utf8');
}

async function blobToBuffer(blob: Blob): Promise<Buffer> {
    return Buffer.from(await blob.arrayBuffer());
}

async function testXlsxFeatures(): Promise<void> {
    const file = path.join(outputDir, 'features.xlsx');
    const writer = new XlsxWriter(file, { compressionLevel: 0 });
    writer.addSheet('Feat');
    writer.writeSheet(ROWS, HEADERS, FEATURES);
    await writer.finalize();

    const buffer = fs.readFileSync(file);
    const parts = unpack(buffer);
    const sheet = partText(parts, 'xl/worksheets/sheet1.xml');
    assert.ok(sheet.includes('<mergeCells count="1"><mergeCell ref="A4:B4"/>'), 'merge cells written');
    assert.ok(sheet.includes('<hyperlink ref="C1" r:id="rId1" tooltip="Example"/>'), 'external hyperlink written');
    assert.ok(sheet.includes('<hyperlink ref="A2" location="Feat!A1"/>'), 'internal hyperlink written');
    assert.ok(sheet.includes('<dataValidation type="list" sqref="C1:C5" allowBlank="1">'), 'validation written');
    assert.ok(sheet.includes('<conditionalFormatting sqref="A2:A3">'), 'conditional formatting written');
    assert.ok(sheet.includes('dxfId="0"'), 'conditional formatting references a dxf');
    assert.ok(sheet.includes('<legacyDrawing r:id="rId3"/>'), 'comment VML linked');
    assert.ok(/<col min="1" max="1" width="30"/.test(sheet), 'fixed width A');
    assert.ok(/<col min="2" max="2" width="40"/.test(sheet), 'fixed width B');
    assert.ok(parts.has('xl/comments/sheet1.xml'), 'comments part exists');
    assert.ok(parts.has('xl/drawings/commentsDrawing1.vml'), 'VML part exists');
    assert.ok(partText(parts, '[Content_Types].xml').includes('/xl/comments/sheet1.xml'), 'comments content type');

    const reader = new XlsxReader();
    await reader.openBuffer(buffer);
    await reader.selectSheet('Feat');
    assert.ok(reader.readRow() !== null, 'header row read');
    assert.ok(reader.readRow() !== null, 'data row read');
    const style = reader.getStyle(0);
    assert.equal(style?.numFmt, '#,##0.00');
    assert.equal(style?.font?.bold, true);
    assert.equal(style?.font?.color, 'FFFF0000');
    assert.equal(style?.fill?.fgColor, 'FFFFFF00');
    assert.equal(style?.border?.top?.style, 'thin');
    assert.equal(style?.alignment?.horizontal, 'center');
    assert.equal(reader.getValue(0), 10);
    assert.equal(reader.getFormula(1), 'SUM(A2:A3)');
    assert.equal(reader.getValue(1), 30);
    const rich = reader.getRichText(2);
    assert.equal(rich?.toString(), 'red plain');
    assert.equal(rich?.runs[1].text, ' plain');
    await reader.close();

    const inlineFile = path.join(outputDir, 'inline-rich.xlsx');
    const inlineWriter = new XlsxWriter(inlineFile, { compressionLevel: 0, useSharedStrings: false });
    inlineWriter.addSheet('Inline');
    inlineWriter.writeSheet([[new RichText([{ text: 'a', style: { font: { italic: true } } }, { text: 'b' }])]], null, false);
    await inlineWriter.finalize();
    const inlineReader = new XlsxReader();
    await inlineReader.open(inlineFile);
    await inlineReader.selectSheet('Inline');
    assert.ok(inlineReader.readRow() !== null, 'inline rich row read');
    assert.equal(inlineReader.getRichText(0)?.toString(), 'ab');
    assert.equal(inlineReader.getRichText(0)?.runs[0].style?.font?.italic, true);
    await inlineReader.close();
}

async function testXlsbFeatures(): Promise<void> {
    const file = path.join(outputDir, 'features.xlsb');
    const writer = new XlsbWriter(file, { compressionLevel: 0 });
    writer.addSheet('Feat');
    writer.writeSheet(XLSB_ROWS, HEADERS, {
        merges: FEATURES.merges,
        hyperlinks: FEATURES.hyperlinks,
        dataValidations: FEATURES.dataValidations,
        columnWidths: FEATURES.columnWidths,
    });
    await writer.finalize();

    const buffer = fs.readFileSync(file);
    const parts = unpack(buffer);
    const sheet = parts.get('xl/worksheets/sheet1.bin');
    assert.ok(sheet, 'XLSB sheet exists');
    // BrtBeginMergeCells / BrtMergeCell / BrtEndMergeCells.
    assert.ok(sheet!.includes(Buffer.from([0xb1, 0x01, 0x04, 0x01, 0x00, 0x00, 0x00])), 'merge records written');
    assert.ok(sheet!.includes(Buffer.from([0xee, 0x03])), 'hyperlink record area written');
    const rels = partText(parts, 'xl/worksheets/_rels/sheet1.bin.rels');
    assert.ok(rels.includes('https://example.com'), 'external hyperlink relationship');
    assert.ok(rels.includes('TargetMode="External"'), 'external target mode');

    const reader = new XlsbReader();
    await reader.openBuffer(buffer);
    await reader.selectSheet('Feat');
    assert.ok(reader.readRow() !== null, 'header row read');
    assert.ok(reader.readRow() !== null, 'data row read');
    const style = reader.getStyle(0);
    assert.equal(style?.numFmt, '#,##0.00');
    assert.equal(style?.font?.bold, true);
    assert.equal(style?.fill?.fgColor, 'FFFFFF00');
    assert.equal(style?.border?.top?.style, 'thin');
    assert.equal(style?.alignment?.horizontal, 'center');
    assert.equal(reader.getValue(0), 10);
    assert.equal(reader.getValue(1), 30);
    await reader.close();

    await assert.rejects(
        (async () => {
            const w = new XlsbWriter(path.join(outputDir, 'unsupported.xlsb'));
            w.addSheet('S');
            w.writeSheet([[1]], ['A'], { comments: FEATURES.comments });
            await w.finalize();
        })(),
        /Cell comments are not supported by XlsbWriter/,
    );
}

async function testStreamingFeatures(): Promise<void> {
    const file = path.join(outputDir, 'stream.xlsx');
    const writer = new XlsxWriter(file, { compressionLevel: 0 });
    writer.startSheet('Stream', 3, HEADERS, {
        merges: ['A4:B4'],
        hyperlinks: [{ ref: 'C1', target: 'https://example.com' }],
        columnWidths: [12, 13, 14],
    });
    for (const row of ROWS) writer.writeRow(row);
    writer.endSheet();
    await writer.finalize();

    const sheet = partText(unpack(fs.readFileSync(file)), 'xl/worksheets/sheet1.xml');
    assert.ok(sheet.includes('<mergeCell ref="A4:B4"/>'), 'streaming merge written');
    assert.ok(sheet.includes('<hyperlink ref="C1" r:id="rId1"/>'), 'streaming hyperlink written');
    assert.ok(/<col min="1" max="1" width="12"/.test(sheet), 'streaming fixed width');

    const xlsbFile = path.join(outputDir, 'stream.xlsb');
    const xlsb = new XlsbWriter(xlsbFile, { compressionLevel: 0 });
    xlsb.startSheet('Stream', 3, HEADERS, {
        merges: ['A4:B4'],
        hyperlinks: [{ ref: 'C1', target: 'https://example.com' }],
        dataValidations: [{ sqref: 'C2:C5', type: 'list', values: ['x', 'y'] }],
        columnWidths: [12, 13, 14],
    });
    for (const row of XLSB_ROWS) xlsb.writeRow(row);
    xlsb.endSheet();
    await xlsb.finalize();

    const xlsbParts = unpack(fs.readFileSync(xlsbFile));
    assert.ok(xlsbParts.get('xl/worksheets/sheet1.bin')!.includes(
        Buffer.from([0xb1, 0x01, 0x04, 0x01, 0x00, 0x00, 0x00]),
    ), 'streaming XLSB merge written');
    assert.ok(partText(xlsbParts, 'xl/worksheets/_rels/sheet1.bin.rels').includes('https://example.com'),
        'streaming XLSB hyperlink relationship');

    // Streaming measures the first 200 rows instead of only the header.
    const sampleRows: CellValue[][] = Array.from({ length: 150 }, (_, row) => [
        row === 149 ? 'WWWWWWWWWW' : 'a',
    ]);
    const delayedFile = path.join(outputDir, 'stream-delayed-width.xlsx');
    const delayed = new XlsxWriter(delayedFile, { compressionLevel: 0 });
    delayed.startSheet('Late', 1, ['A']);
    for (const row of sampleRows) delayed.writeRow(row);
    delayed.endSheet();
    await delayed.finalize();
    const delayedSheet = partText(unpack(fs.readFileSync(delayedFile)), 'xl/worksheets/sheet1.xml');
    const delayedWidth = Number(/<col[^>]*width="([^"]+)"/.exec(delayedSheet)![1]);
    assert.ok(delayedWidth > autofitFileWidth('a'),
        `streaming width measured row 150, got ${delayedWidth}`);

    // Node and browser streaming builds stay byte-identical.
    const browserDelayed = new BrowserXlsxWriter();
    browserDelayed.startSheet('Late', 1, ['A']);
    for (const row of sampleRows) browserDelayed.writeRow(row);
    browserDelayed.endSheet();
    const browserDelayedParts = unpack(await blobToBuffer(browserDelayed.finalize()));
    assert.ok(browserDelayedParts.get('xl/worksheets/sheet1.xml')!
        .equals(unpack(fs.readFileSync(delayedFile)).get('xl/worksheets/sheet1.xml')!),
    'browser streaming measures the same rows as Node');

    const nodeStreamXlsb = path.join(outputDir, 'stream-delayed.xlsb');
    const nodeXlsb = new XlsbWriter(nodeStreamXlsb, { compressionLevel: 0 });
    nodeXlsb.startSheet('Late', 1, ['A']);
    for (const row of sampleRows) nodeXlsb.writeRow(row);
    nodeXlsb.endSheet();
    await nodeXlsb.finalize();

    const browserXlsb = new BrowserXlsbWriter();
    browserXlsb.startSheet('Late', 1, ['A']);
    for (const row of sampleRows) browserXlsb.writeRow(row);
    browserXlsb.endSheet();
    const browserXlsbParts = unpack(await blobToBuffer(browserXlsb.finalize()));
    assert.ok(browserXlsbParts.get('xl/worksheets/sheet1.bin')!
        .equals(unpack(fs.readFileSync(nodeStreamXlsb)).get('xl/worksheets/sheet1.bin')!),
    'browser XLSB streaming matches Node');
}

async function testXlsxParity(): Promise<void> {
    const nodeFile = path.join(outputDir, 'xlsx-parity-node.xlsx');
    const node = new XlsxWriter(nodeFile, { compressionLevel: 0 });
    node.addSheet('Feat');
    node.writeSheet(ROWS, HEADERS, FEATURES);
    await node.finalize();

    const browser = new BrowserXlsxWriter();
    browser.addSheet('Feat');
    browser.writeSheet(ROWS, HEADERS, FEATURES);
    const browserBuffer = await blobToBuffer(browser.finalize());

    const left = unpack(fs.readFileSync(nodeFile));
    const right = unpack(browserBuffer);
    assert.deepEqual([...right.keys()].sort(), [...left.keys()].sort(), 'XLSX parts match');
    for (const [name, data] of left) {
        assert.ok(right.get(name)?.equals(data), `XLSX part ${name} is byte-identical`);
    }
}

async function testParity(): Promise<void> {
    const nodeFile = path.join(outputDir, 'parity-node.xlsb');
    const node = new XlsbWriter(nodeFile, { compressionLevel: 0 });
    node.addSheet('Feat');
    node.writeSheet(XLSB_ROWS, HEADERS, {
        merges: FEATURES.merges,
        hyperlinks: FEATURES.hyperlinks,
        dataValidations: FEATURES.dataValidations,
        columnWidths: FEATURES.columnWidths,
    });
    await node.finalize();

    const browser = new BrowserXlsbWriter();
    browser.addSheet('Feat');
    browser.writeSheet(XLSB_ROWS, HEADERS, {
        merges: FEATURES.merges,
        hyperlinks: FEATURES.hyperlinks,
        dataValidations: FEATURES.dataValidations,
        columnWidths: FEATURES.columnWidths,
    });
    const browserBuffer = await blobToBuffer(browser.finalize());

    const left = unpack(fs.readFileSync(nodeFile));
    const right = unpack(browserBuffer);
    assert.deepEqual([...right.keys()].sort(), [...left.keys()].sort(), 'XLSB parts match');
    for (const [name, data] of left) {
        assert.ok(right.get(name)?.equals(data), `XLSB part ${name} is byte-identical`);
    }
}

async function testFormulaCacheTypes(): Promise<void> {
    const rows: CellValue[][] = [[
        { formula: '"ok"', value: 'ok' } as FormulaCell,
        { formula: 'TRUE', value: true } as FormulaCell,
        { formula: 'NA()', value: new ExcelError('#N/A') } as FormulaCell,
        { formula: 'DATE(2024,1,2)', value: new Date(Date.UTC(2024, 0, 2)) } as FormulaCell,
    ]];
    for (const format of ['xlsx', 'xlsb'] as const) {
        const file = path.join(outputDir, `cache-${format}.${format}`);
        const writer = format === 'xlsx' ? new XlsxWriter(file, { compressionLevel: 0 }) : new XlsbWriter(file, { compressionLevel: 0 });
        writer.addSheet('Cache');
        writer.writeSheet(rows, null, false);
        await writer.finalize();

        const reader = format === 'xlsx' ? new XlsxReader() : new XlsbReader();
        await reader.open(file);
        await reader.selectSheet('Cache');
        assert.ok(reader.readRow() !== null, `${format} cache row read`);
        assert.equal(reader.getValue(0), 'ok', `${format} cached string`);
        assert.equal(reader.getValue(1), true, `${format} cached boolean`);
        assert.equal(reader.getValue(2), '#N/A', `${format} cached error`);
        assert.ok(reader.getValue(3) instanceof Date, `${format} cached date`);
        await reader.close();
    }
}

async function testAutofitWidths(): Promise<void> {
    assert.equal(autofitWidth('A'), DEFAULT_COLUMN_WIDTH, 'short text uses the 8.43 minimum');
    assert.ok(Math.abs(autofitFileWidth('A') - (DEFAULT_COLUMN_WIDTH + COLUMN_WIDTH_PADDING)) < 1e-9,
        'autofit adds the 5/7 cell padding');
    assert.equal(textDisplayUnits('ABC'), 3.8);
    assert.ok(Math.abs(textDisplayUnits('日本語') - 6.6) < 1e-9,
        'CJK characters count as 2.2 units');
    assert.equal(autofitWidth('x'.repeat(250)), MAX_AUTOFIT_WIDTH,
        'autofit is capped at 200');
    assert.equal(autofitFileWidth('x'.repeat(250)), MAX_AUTOFIT_WIDTH,
        'the 200 cap applies to the stored width including padding');
    assert.ok(autofitWidth('WWWW', { name: 'Arial' }) > autofitWidth('WWWW'),
        'wide fonts use the wider conservative cost table');

    // Unicode outliers measured with xlspy's widest font advances.
    assert.ok(autofitWidth('Ǆ'.repeat(30), { name: 'Aptos', size: 14 }) >= 115.71,
        'wide Latin Extended letters are not under-measured');
    assert.ok(autofitWidth('Ѿ'.repeat(30), { name: 'Arial Narrow' }) >= 93.29,
        'wide Cyrillic letters are not under-measured');
    assert.ok(autofitWidth('‱'.repeat(30), { name: 'Tahoma', bold: true }) >= 142.43,
        'wide symbol characters are not under-measured');

    // Indent: Excel reserves 9 px per level, twice for centred text.
    const plainWidth = [-1];
    updateColWidthsFromRows(plainWidth, [[{ value: 'Description', style: { font: { name: 'Calibri' } } } as StyledCell]]);
    const indentedWidth = [-1];
    updateColWidthsFromRows(indentedWidth, [[{
        value: 'Description',
        style: { font: { name: 'Calibri' }, alignment: { indent: 5 } },
    } as StyledCell]]);
    assert.ok(indentedWidth[0] > plainWidth[0], 'alignment indent widens the column');

    // Wrap text measures the longest line, not the sum of all lines.
    const lines = Array(20).fill('abcdefghij').join('\n');
    const wrapped = [-1];
    updateColWidthsFromRows(wrapped, [[{
        value: lines,
        style: { font: { name: 'Calibri' }, alignment: { wrapText: true } },
    } as StyledCell]]);
    assert.ok(wrapped[0] < 20, `wrapped text stays near one line, got ${wrapped[0]}`);

    // Negative number-format sections carry their own literal text.
    const negative = [-1];
    updateColWidthsFromRows(negative, [[{
        value: -1,
        style: { numFmt: '0.00;[Red]"A very long negative literal: "#,##0.00' },
    } as StyledCell]]);
    assert.ok(negative[0] >= 29.43, `negative section literal is measured, got ${negative[0]}`);

    // General numbers use Excel's 11-character display instead of the full
    // JavaScript decimal expansion.
    const general = [-1];
    updateColWidthsFromRows(general, [[{
        value: 1.2345678901234567,
        style: { font: { name: 'Calibri' } },
    } as StyledCell]]);
    assert.ok(general[0] < 20, `General long decimal stays compact, got ${general[0]}`);

    // The default sample window is 200 rows (xlspy measures up to 1000).
    const sampleRows: CellValue[][] = Array.from({ length: 150 }, (_, row) => [
        row === 149 ? 'WWWWWWWWWW' : 'a',
    ]);
    const sampled = [-1];
    updateColWidthsFromRows(sampled, sampleRows);
    assert.equal(AUTOFIT_SAMPLE_ROWS, 200);
    assert.ok(sampled[0] > autofitFileWidth('a'), 'wide value at row 150 is measured');
    const oldWindow = [-1];
    updateColWidthsFromRows(oldWindow, sampleRows, 100);
    assert.ok(oldWindow[0] < sampled[0], 'an explicit 100-row window still stops early');

    const file = path.join(outputDir, 'autofit.xlsx');
    const writer = new XlsxWriter(file, { compressionLevel: 0 });
    writer.addSheet('Auto');
    writer.writeSheet([], ['A', 'Description']);
    await writer.finalize();
    const sheet = partText(unpack(fs.readFileSync(file)), 'xl/worksheets/sheet1.xml');
    assert.ok(sheet.includes('width="11.54"'),
        'XLSX header width includes the minimum, padding and filter button');
    assert.ok(sheet.includes('width="20.87"'),
        'XLSX writes the padded bold autofit width');
    assert.ok(AUTOFILTER_BUTTON_WIDTH >= 16 / 7,
        'filter button reservation is not smaller than xlspy');

    const noFilterFile = path.join(outputDir, 'autofit-no-filter.xlsx');
    const noFilterWriter = new XlsxWriter(noFilterFile, { compressionLevel: 0 });
    noFilterWriter.addSheet('NoFilter');
    noFilterWriter.writeSheet([], ['A', 'Description'], { doAutofilter: false });
    await noFilterWriter.finalize();
    const noFilterSheet = partText(unpack(fs.readFileSync(noFilterFile)), 'xl/worksheets/sheet1.xml');
    assert.ok(noFilterSheet.includes('width="9.14"'),
        'without a filter button the header keeps the minimum width');
    assert.ok(noFilterSheet.includes('width="18.47"'),
        'without a filter button the padded bold width is smaller');
    assert.ok(!noFilterSheet.includes('<autoFilter'),
        'doAutofilter: false in the options object is honoured');

    // Dates with a time keep the time visible (numFmt 22, style 2).
    const dateFile = path.join(outputDir, 'date-styles.xlsx');
    const dateWriter = new XlsxWriter(dateFile, { compressionLevel: 0 });
    dateWriter.addSheet('Dates');
    dateWriter.writeSheet([
        [new Date(Date.UTC(2024, 0, 2))],
        [new Date(Date.UTC(2024, 0, 2, 15, 30))],
    ]);
    await dateWriter.finalize();
    const dateSheet = partText(unpack(fs.readFileSync(dateFile)), 'xl/worksheets/sheet1.xml');
    assert.ok(dateSheet.includes('<c s="1">'), 'date-only cells use the date style');
    assert.ok(dateSheet.includes('<c s="2">'), 'datetime cells use the datetime style');

    const badWidthWriter = new XlsxWriter(path.join(outputDir, 'bad-width.xlsx'));
    badWidthWriter.addSheet('Bad');
    assert.throws(
        () => badWidthWriter.writeSheet([], ['A'], { columnWidths: [0] }),
        /0\.5 and 255/,
        'explicit column widths are validated',
    );
}

async function main(): Promise<void> {
    try {
        await testXlsxFeatures();
        await testXlsbFeatures();
        await testStreamingFeatures();
        await testXlsxParity();
        await testParity();
        await testFormulaCacheTypes();
        await testAutofitWidths();
        console.log('Stage 2 feature tests passed');
    } finally {
        fs.rmSync(outputDir, { recursive: true, force: true });
    }
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
