import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxReader } from '../../src/XlsxReader';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-xlsx-refs-'));

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

// The r attribute on <c> is optional (ECMA-376). Writers such as xlspy omit it,
// so a cell without r must land in the column after the previous cell.
function createXlsxWithoutCellRefs(filePath: string): void {
    const zip = new AdmZip();
    const addText = (name: string, text: string) => zip.addFile(name, Buffer.from(text, 'utf8'));
    addText('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    addText('xl/workbook.xml',
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    addText('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    // Row 1: three cells without r (A, B, C), then an explicit E1, then one without r (F1).
    // Row 2: all cells without r, plus an explicit A2 (a reference resets the position).
    addText('xl/worksheets/sheet1.xml',
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c t="inlineStr"><is><t>a</t></is></c><c><v>2</v></c><c t="b"><v>1</v></c><c r="E1" t="inlineStr"><is><t>e</t></is></c><c><v>6</v></c></row>' +
        '<row r="2"><c r="A2"><v>10</v></c><c><v>11</v></c><c><v>12</v></c></row>' +
        '</sheetData></worksheet>');
    zip.writeZip(filePath);
}

async function testCellsWithoutRefs(): Promise<void> {
    const filePath = path.join(outputDir, 'no-cell-refs.xlsx');
    createXlsxWithoutCellRefs(filePath);

    const reader = new XlsxReader();
    await reader.open(filePath);

    assert(await reader.read(), 'first row is readable');
    assert(reader.getValue(0) === 'a', `A1 without r is placed in column A, got ${reader.getValue(0)}`);
    assert(reader.getValue(1) === 2, `B1 without r is placed in column B, got ${reader.getValue(1)}`);
    assert(reader.getValue(2) === true, `C1 without r is placed in column C, got ${reader.getValue(2)}`);
    assert(reader.getValue(3) == null, `skipped D1 stays empty, got ${reader.getValue(3)}`);
    assert(reader.getValue(4) === 'e', `explicit E1 is kept, got ${reader.getValue(4)}`);
    assert(reader.getValue(5) === 6, `F1 without r follows E1, got ${reader.getValue(5)}`);

    assert(await reader.read(), 'second row is readable');
    assert(reader.getValue(0) === 10, `explicit A2 is kept, got ${reader.getValue(0)}`);
    assert(reader.getValue(1) === 11, `B2 without r follows A2, got ${reader.getValue(1)}`);
    assert(reader.getValue(2) === 12, `C2 without r follows B2, got ${reader.getValue(2)}`);

    assert(!(await reader.read()), 'no more rows after the second row');
    await reader.close();
}

// Serial dates with a time part (as written by xlspy) must map to the exact millisecond.
// Excel serial 44927 is 2023-01-01; the fraction is seconds / 86400.
function createXlsxWithDateTimes(filePath: string, seconds: number[]): void {
    const zip = new AdmZip();
    const addText = (name: string, text: string) => zip.addFile(name, Buffer.from(text, 'utf8'));
    addText('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    addText('xl/workbook.xml',
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    addText('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
    // Style 1 uses built-in number format 22 (m/d/yy h:mm), which the reader treats as a date.
    addText('xl/styles.xml',
        '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="22" applyNumberFormat="1"/></cellXfs></styleSheet>');
    const cells = seconds
        .map((s, i) => `<c r="${String.fromCharCode(65 + i)}1" s="1"><v>${44927 + (12 * 3600 + 30 * 60 + s) / 86400}</v></c>`)
        .join('');
    addText('xl/worksheets/sheet1.xml',
        `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`);
    zip.writeZip(filePath);
}

async function testDateTimeMilliseconds(): Promise<void> {
    const seconds = Array.from({ length: 26 }, (_, i) => i * 2 + 1);
    const filePath = path.join(outputDir, 'date-times.xlsx');
    createXlsxWithDateTimes(filePath, seconds);

    const reader = new XlsxReader();
    await reader.open(filePath);
    assert(await reader.read(), 'date row is readable');
    seconds.forEach((s, i) => {
        const value = reader.getValue(i);
        const expected = Date.UTC(2023, 0, 1, 12, 30, s);
        assert(value instanceof Date, `column ${i} is a Date, got ${value}`);
        assert((value as Date).getTime() === expected,
            `12:30:${String(s).padStart(2, '0')} maps to the exact millisecond, got ${(value as Date).toISOString()}`);
    });
    await reader.close();
}

async function main(): Promise<void> {
    await testCellsWithoutRefs();
    await testDateTimeMilliseconds();
    console.log('XLSX reader cell reference tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
