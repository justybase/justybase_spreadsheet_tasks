import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { oaDateToTime, timeToOaDate } from '../../src/writerHelpers';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsbUpdater } from '../../src/XlsbUpdater';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-oa-leap-'));
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

function testSerialHelpers(): void {
    // The 1900 system counts a fictitious 1900-02-29 (serial 60); serials
    // 1-59 denote 1900-01-01..1900-02-28 (Excel/xlspy/openpyxl parity).
    assert(oaDateToTime(1, EPOCH_1900) === Date.UTC(1900, 0, 1), 'serial 1 -> 1900-01-01');
    assert(oaDateToTime(59, EPOCH_1900) === Date.UTC(1900, 1, 28), 'serial 59 -> 1900-02-28');
    assert(oaDateToTime(60, EPOCH_1900) === Date.UTC(1900, 1, 28), 'serial 60 maps to 1900-02-28');
    assert(oaDateToTime(61, EPOCH_1900) === Date.UTC(1900, 2, 1), 'serial 61 -> 1900-03-01');
    assert(timeToOaDate(Date.UTC(1900, 0, 1), EPOCH_1900) === 1, '1900-01-01 -> serial 1');
    assert(timeToOaDate(Date.UTC(1900, 1, 28), EPOCH_1900) === 59, '1900-02-28 -> serial 59');
    assert(timeToOaDate(Date.UTC(1900, 2, 1), EPOCH_1900) === 61, '1900-03-01 -> serial 61');
    assert(timeToOaDate(Date.UTC(2024, 0, 15), EPOCH_1900) === 45306, '2024-01-15 -> serial 45306');
    // The 1904 system has no leap-day gap.
    assert(oaDateToTime(0, EPOCH_1904) === EPOCH_1904, '1904 serial 0 -> epoch');
    assert(timeToOaDate(EPOCH_1904, EPOCH_1904) === 0, '1904 epoch -> serial 0');
}

function createXlsxWithEarlySerials(filePath: string): void {
    const zip = new AdmZip();
    const addText = (name: string, text: string) => zip.addFile(name, Buffer.from(text, 'utf8'));
    addText('[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    addText('xl/workbook.xml',
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    addText('xl/_rels/workbook.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    addText('xl/styles.xml',
        '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>');
    addText('xl/worksheets/sheet1.xml',
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
        '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>59</v></c><c r="C1" s="1"><v>61</v></c></row>' +
        '</sheetData></worksheet>');
    zip.writeZip(filePath);
}

async function testXlsxReaderEarlySerials(): Promise<void> {
    const f = path.join(outputDir, 'early.xlsx');
    createXlsxWithEarlySerials(f);
    const r = new XlsxReader();
    await r.open(f);
    assert(await r.read(), 'row readable');
    const a = r.getValue(0);
    const b = r.getValue(1);
    const c = r.getValue(2);
    assert(a instanceof Date && a.getTime() === Date.UTC(1900, 0, 1), `serial 1 reads 1900-01-01, got ${a}`);
    assert(b instanceof Date && b.getTime() === Date.UTC(1900, 1, 28), `serial 59 reads 1900-02-28, got ${b}`);
    assert(c instanceof Date && c.getTime() === Date.UTC(1900, 2, 1), `serial 61 reads 1900-03-01, got ${c}`);
    await r.close();
}

async function testXlsbUpdaterKeepsCustomDateStyle(): Promise<void> {
    // Dates land in a column whose old cells were all general-styled, so the
    // updater must fall back to a date xf. The writer emits dates with custom
    // numFmts (164/166), which the fallback has to recognise.
    const src = path.join(outputDir, 'dates.xlsb');
    const w = new XlsbWriter(src);
    w.addSheet('data1');
    w.writeSheet([[1], [2]], ['N'], false);
    await w.finalize();

    const out = path.join(outputDir, 'dates-out.xlsb');
    const u = new XlsbUpdater(src);
    u.replaceSheetData('data1', [[new Date(Date.UTC(2025, 5, 30, 12, 0, 0))]], { headers: ['N'] });
    u.save(out);

    const r = new XlsbReader();
    await r.open(out);
    assert(await r.read(), 'header row readable');
    assert(await r.read(), 'data row readable');
    const v = r.getValue(0);
    assert(v instanceof Date, `updated date cell keeps a date style, got ${v}`);
    assert((v as Date).getTime() === Date.UTC(2025, 5, 30, 12, 0, 0), `date value preserved, got ${v}`);
    await r.close();
}

async function main(): Promise<void> {
    testSerialHelpers();
    await testXlsxReaderEarlySerials();
    await testXlsbUpdaterKeepsCustomDateStyle();
    console.log('1900 leap-day OA tests passed');
}

main().catch((e) => { console.error(e); process.exitCode = 1; });
