/**
 * Shared-core parity tests.
 *
 * The Node writers and the browser writers must emit byte-identical unpacked
 * package parts for the same data (ZIP compression differs, the contents must
 * not). This is the acceptance gate for the shared writer core.
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { BrowserXlsxWriter } from '../../src/browser/BrowserXlsxWriter';
import { BrowserXlsbWriter } from '../../src/browser/BrowserXlsbWriter';
import { CellValue, ExcelError, FormattedCell } from '../../src/Formats';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-core-parity-'));

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

const HEADERS = ['Name', 'Count', 'Ratio', 'When', 'Flag', 'Note'];
const ROWS: CellValue[][] = [
    ['alpha', 1, 1.5, new Date(Date.UTC(2024, 0, 2, 3, 4, 5)), true, new ExcelError('#DIV/0!')],
    ['ąęćśńźół 日本語', -42, 0.25, new Date(Date.UTC(1999, 11, 31, 23, 59, 59)), false, { value: 12.345, format: '0.00' } as FormattedCell],
    ['  spaced  ', 0, -0.5, new Date(Date.UTC(2020, 5, 15, 12, 0, 0)), true, null],
    ['_x0041_ literal', 1000, 1e-9, new Date(Date.UTC(2023, 2, 1)), false, 'tail'],
];

async function blobToBuffer(blob: Blob): Promise<Buffer> {
    return Buffer.from(await blob.arrayBuffer());
}

function unpack(buffer: Buffer): Map<string, Buffer> {
    const zip = new AdmZip(buffer);
    const parts = new Map<string, Buffer>();
    for (const entry of zip.getEntries()) {
        if (!entry.isDirectory) parts.set(entry.entryName, entry.getData());
    }
    return parts;
}

function compareParts(a: Buffer, b: Buffer, format: string): void {
    const left = unpack(a);
    const right = unpack(b);
    const names = [...new Set([...left.keys(), ...right.keys()])].sort();
    for (const name of names) {
        const l = left.get(name);
        const r = right.get(name);
        assert(l !== undefined && r !== undefined, `${format}: part ${name} exists on both sides`);
        assert(l!.equals(r!), `${format}: part ${name} is byte-identical (${l!.length} vs ${r!.length})`);
    }
}

async function writeXlsxNode(inline: boolean): Promise<Buffer> {
    const file = path.join(outputDir, `node-${inline ? 'inline' : 'shared'}.xlsx`);
    const writer = new XlsxWriter(file, { compressionLevel: 0, useSharedStrings: !inline });
    writer.addSheet('Data');
    writer.writeSheet(ROWS, HEADERS, true);
    writer.startSheet('Stream', 2, ['A', 'B'], { doAutofilter: true });
    writer.writeRow(['x', 1]);
    writer.writeRow(['y', 2]);
    writer.endSheet();
    await writer.finalize();
    return fs.readFileSync(file);
}

async function writeXlsxBrowser(inline: boolean): Promise<Buffer> {
    const writer = new BrowserXlsxWriter({ useSharedStrings: !inline });
    writer.addSheet('Data');
    writer.writeSheet(ROWS, HEADERS, { doAutofilter: true });
    writer.startSheet('Stream', 2, ['A', 'B'], { doAutofilter: true });
    writer.writeRow(['x', 1]);
    writer.writeRow(['y', 2]);
    writer.endSheet();
    return blobToBuffer(writer.finalize());
}

async function writeXlsbNode(inline: boolean): Promise<Buffer> {
    const file = path.join(outputDir, `node-${inline ? 'inline' : 'shared'}.xlsb`);
    const writer = new XlsbWriter(file, { compressionLevel: 0, useSharedStrings: !inline });
    writer.addSheet('Data');
    writer.writeSheet(ROWS, HEADERS, true);
    writer.startSheet('Stream', 2, ['A', 'B'], { doAutofilter: true });
    writer.writeRow(['x', 1]);
    writer.writeRow(['y', 2]);
    writer.endSheet();
    await writer.finalize();
    return fs.readFileSync(file);
}

async function writeXlsbBrowser(inline: boolean): Promise<Buffer> {
    const writer = new BrowserXlsbWriter({ useSharedStrings: !inline });
    writer.addSheet('Data');
    writer.writeSheet(ROWS, HEADERS, { doAutofilter: true });
    writer.startSheet('Stream', 2, ['A', 'B'], { doAutofilter: true });
    writer.writeRow(['x', 1]);
    writer.writeRow(['y', 2]);
    writer.endSheet();
    return blobToBuffer(writer.finalize());
}

async function readRows(reader: XlsxReader | XlsbReader, file: string, sheet: string): Promise<CellValue[][]> {
    await reader.open(file);
    await reader.selectSheet(sheet);
    return reader.readAll();
}

async function testXlsxParity(): Promise<void> {
    for (const inline of [false, true]) {
        const node = await writeXlsxNode(inline);
        const browser = await writeXlsxBrowser(inline);
        compareParts(node, browser, `xlsx ${inline ? 'inline' : 'shared'}`);
    }

    const file = path.join(outputDir, 'node-shared.xlsx');
    const rows = await readRows(new XlsxReader(), file, 'Data');
    assert(rows.length === ROWS.length + 1, `xlsx read returns header + ${ROWS.length} data rows, got ${rows.length}`);
    assert(rows[0][0] === 'Name', 'xlsx header round-trips');
    assert(rows[1][0] === 'alpha', 'xlsx string round-trips');
    assert(rows[2][0] === 'ąęćśńźół 日本語', 'xlsx unicode round-trips');
    assert(rows[1][5] === '#DIV/0!', 'xlsx error round-trips');
    assert(rows[2][5] === 12.345, 'xlsx formatted number round-trips');
    assert(rows[1][3] instanceof Date, 'xlsx date round-trips');

    const inlineRows = await readRows(new XlsxReader(), path.join(outputDir, 'node-inline.xlsx'), 'Data');
    assert(inlineRows[1][0] === 'alpha', 'xlsx inline string round-trips');
    assert(inlineRows[2][0] === 'ąęćśńźół 日本語', 'xlsx inline unicode round-trips');

    const zip = new AdmZip(fs.readFileSync(path.join(outputDir, 'node-inline.xlsx')));
    const sheet = zip.readAsText('xl/worksheets/sheet1.xml');
    assert(sheet.includes('t="inlineStr"'), 'xlsx inline mode writes inlineStr cells');
}

async function testXlsbParity(): Promise<void> {
    for (const inline of [false, true]) {
        const node = await writeXlsbNode(inline);
        const browser = await writeXlsbBrowser(inline);
        compareParts(node, browser, `xlsb ${inline ? 'inline' : 'shared'}`);
    }

    const zip = new AdmZip(fs.readFileSync(path.join(outputDir, 'node-shared.xlsb')));
    const sheet = zip.getEntry('xl/worksheets/sheet1.bin')!.getData();
    const dim = sheet.indexOf(Buffer.from([0x94, 0x01, 0x10]));
    assert(dim !== -1, 'xlsb sheet has BrtWsDim');
    assert(sheet.readInt32LE(dim + 3) === 0, 'xlsb dimension first row is 0');
    assert(sheet.readInt32LE(dim + 7) === ROWS.length, `xlsb dimension last row is ${ROWS.length}, got ${sheet.readInt32LE(dim + 7)}`);
    assert(sheet.readInt32LE(dim + 15) === HEADERS.length - 1, 'xlsb dimension last column is patched');

    const rows = await readRows(new XlsbReader(), path.join(outputDir, 'node-shared.xlsb'), 'Data');
    assert(rows.length === ROWS.length + 1, `xlsb read returns header + ${ROWS.length} data rows, got ${rows.length}`);
    assert(rows[0][0] === 'Name', 'xlsb header round-trips');
    assert(rows[1][0] === 'alpha', 'xlsb string round-trips');
    assert(rows[2][0] === 'ąęćśńźół 日本語', 'xlsb unicode round-trips');
    assert(rows[1][5] === '#DIV/0!', 'xlsb error round-trips');
    assert(rows[2][5] === 12.345, 'xlsb formatted number round-trips');
    assert(rows[1][3] instanceof Date, 'xlsb date round-trips');

    const inlineRows = await readRows(new XlsbReader(), path.join(outputDir, 'node-inline.xlsb'), 'Data');
    assert(inlineRows[1][0] === 'alpha', 'xlsb inline string round-trips');
    assert(inlineRows[2][0] === 'ąęćśńźół 日本語', 'xlsb inline unicode round-trips');
}

async function testLargeXlsbSheet(): Promise<void> {
    const rows = 3000;
    const nodeFile = path.join(outputDir, 'big.xlsb');
    const node = new XlsbWriter(nodeFile, { compressionLevel: 0 });
    node.startSheet('Big', 3);
    for (let r = 0; r < rows; r++) node.writeRow([r, `row-${r}`, r / 10]);
    node.endSheet();
    await node.finalize();

    const browser = new BrowserXlsbWriter();
    browser.startSheet('Big', 3);
    for (let r = 0; r < rows; r++) browser.writeRow([r, `row-${r}`, r / 10]);
    browser.endSheet();
    const browserBytes = await blobToBuffer(browser.finalize());

    compareParts(fs.readFileSync(nodeFile), browserBytes, 'xlsb large sheet');

    const zip = new AdmZip(fs.readFileSync(nodeFile));
    const sheet = zip.getEntry('xl/worksheets/sheet1.bin')!.getData();
    const dim = sheet.indexOf(Buffer.from([0x94, 0x01, 0x10]));
    assert(dim !== -1, 'large xlsb sheet has BrtWsDim');
    assert(sheet.readInt32LE(dim + 7) === rows - 1,
        `large xlsb dimension last row is ${rows - 1}, got ${sheet.readInt32LE(dim + 7)}`);
}

async function testBufferOpen(): Promise<void> {
    const xlsxBytes = fs.readFileSync(path.join(outputDir, 'node-shared.xlsx'));
    const xlsx = new XlsxReader();
    await xlsx.openBuffer(xlsxBytes);
    assert(xlsx.getSheetNames()[0] === 'Data', 'xlsx openBuffer keeps sheet names');
    await xlsx.selectSheet('Data');
    const xlsxRows = await xlsx.readAll();
    assert(xlsxRows[1][0] === 'alpha', 'xlsx openBuffer reads values');
    await xlsx.close();

    const xlsbBytes = fs.readFileSync(path.join(outputDir, 'node-shared.xlsb'));
    const xlsb = new XlsbReader();
    await xlsb.openBuffer(xlsbBytes);
    await xlsb.selectSheet('Data');
    const streamed: unknown[][] = [];
    for await (const row of xlsb.rows()) streamed.push(row);
    assert(streamed.length === ROWS.length + 1, 'xlsb rows() yields the header and every data row');
    assert(streamed[1][0] === 'alpha', 'xlsb rows() reads values');
    await xlsb.close();
}

async function main(): Promise<void> {
    await testXlsxParity();
    await testXlsbParity();
    await testLargeXlsbSheet();
    await testBufferOpen();
    console.log('Writer core parity tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
