import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsbWriter } from '../../src/XlsbWriter';
import { readRecord, Biff12Record } from '../../src/biff12Utils';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-row-limit-'));
const HUGE_ROW = 2147483647;
// Reads beyond this bound mean the reader is spinning through phantom gap rows.
const READ_BOUND = 100;

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

function createXlsxWithHugeRowNumber(filePath: string): void {
    const zip = new AdmZip();
    const addText = (name: string, text: string) => zip.addFile(name, Buffer.from(text, 'utf8'));
    addText('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    addText('xl/workbook.xml',
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    addText('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    addText('xl/worksheets/sheet1.xml',
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        `<row r="1"><c r="A1"><v>1</v></c></row><row r="${HUGE_ROW}"><c r="A${HUGE_ROW}"><v>2</v></c></row>` +
        '</sheetData></worksheet>');
    zip.writeZip(filePath);
}

async function testXlsxHugeRowNumber(): Promise<void> {
    const f = path.join(outputDir, 'huge-row.xlsx');
    createXlsxWithHugeRowNumber(f);

    const reader = new XlsxReader();
    await reader.open(f);
    const values: unknown[] = [];
    let reads = 0;
    while (reads < READ_BOUND && (await reader.read())) {
        reads++;
        values.push(reader.getValue(0));
    }
    await reader.close();

    assert(reads === 2, `XLSX: a row number above the Excel limit must not expand into phantom rows, got ${reads} reads`);
    assert(values[0] === 1 && values[1] === 2, `XLSX: values kept in order, got ${JSON.stringify(values)}`);
}

async function testXlsbHugeRowIndex(): Promise<void> {
    const f = path.join(outputDir, 'huge-row.xlsb');
    await createXlsbWithHugeRowIndexAsync(f);

    const reader = new XlsbReader();
    await reader.open(f);
    const values: unknown[] = [];
    let reads = 0;
    while (reads < READ_BOUND && (await reader.read())) {
        reads++;
        values.push(reader.getValue(0));
    }
    await reader.close();

    assert(reads === 2, `XLSB: a row index above the Excel limit must not expand into phantom rows, got ${reads} reads`);
    assert(values[0] === 1 && values[1] === 2, `XLSB: values kept in order, got ${JSON.stringify(values)}`);
}

async function createXlsbWithHugeRowIndexAsync(filePath: string): Promise<void> {
    const source = path.join(outputDir, 'base-for-huge-row.xlsb');
    const w = new XlsbWriter(source);
    w.addSheet('data1');
    w.writeSheet([[1], [2]], null, false);
    await w.finalize();

    const zip = new AdmZip(source);
    const sheet = zip.getEntry('xl/worksheets/sheet1.bin')!.getData();
    const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    const rowHeaders: number[] = [];
    let pos = 0;
    while (readRecord(sheet, pos, rec)) {
        if (rec.id === 0x00) rowHeaders.push(rec.dataStart);
        pos = rec.dataEnd;
    }
    assert(rowHeaders.length === 2, `expected two BrtRowHdr records, got ${rowHeaders.length}`);
    sheet.writeInt32LE(HUGE_ROW, rowHeaders[1]);
    zip.updateFile('xl/worksheets/sheet1.bin', sheet);
    zip.writeZip(filePath);
}

async function main(): Promise<void> {
    await testXlsxHugeRowNumber();
    await testXlsbHugeRowIndex();
    console.log('Reader row limit tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
