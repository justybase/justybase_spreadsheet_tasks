import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxWriter } from '../../src/XlsxWriter';
import { escapeXlsxText, escapeXmlText, unescapeXml, unescapeXmlText, parseSharedStringsXml, parseTagAttributes, resolveRelTarget } from '../../src/xmlUtils';
import { createMinimalXlsx, mkTempDir } from '../helpers/minimalXlsx';

const outputDir = mkTempDir('spreadsheet-reader-xml-');

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

async function testUnescape(): Promise<void> {
    assert(unescapeXml('&amp;lt;') === '&lt;', 'single-pass amp-lt');
    assert(unescapeXml('&lt;') === '<', 'lt');
    assert(unescapeXml('&#65;') === 'A', 'decimal');
    assert(unescapeXml('&#x41;') === 'A', 'hex');
    assert(unescapeXml('&amp;amp;') === '&amp;', 'double amp single pass');
}

async function testParseTagAttributes(): Promise<void> {
    const attrs = parseTagAttributes('<Relationship Target="worksheets/sheet1.xml" Id="rId1" Type="x"/>');
    assert(attrs['Id'] === 'rId1', 'order-independent Id');
    assert(attrs['Target'] === 'worksheets/sheet1.xml', 'order-independent Target');
    const single = parseTagAttributes("<sheet name='a b' r:id='rId2'/>");
    assert(single['name'] === 'a b', 'single quotes');
    assert(single['r:id'] === 'rId2' || single['id'] === 'rId2', 'r:id prefix');
    assert(resolveRelTarget('/xl/worksheets/sheet1.xml') === 'xl/worksheets/sheet1.xml', 'absolute target');
    assert(resolveRelTarget('worksheets/sheet1.xml') === 'xl/worksheets/sheet1.xml', 'relative target');
    assert(resolveRelTarget('xl/worksheets/sheet1.xml') === 'xl/worksheets/sheet1.xml', 'xl target');
}

async function testAbsolutePathAndAttrOrder(): Promise<void> {
    const f = path.join(outputDir, 'abs-order.xlsx');
    createMinimalXlsx(f, {
        workbookXml: '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet r:id="rId1" sheetId="1" name="data1"/></sheets><workbookPr date1904="1"/></workbook>',
        relsXml: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Target="/xl/worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId1"/></Relationships>',
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
    });
    const r = new XlsxReader();
    await r.open(f);
    assert(r.getSheetNames()[0] === 'data1', 'abs path + reordered attrs sheet found');
    assert(await r.read(), 'row readable');
    assert(r.getValue(0) === 1, 'value ok');
    // 1904 epoch: OA date 0 should be 1904-01-01? Check epoch set.
    assert((r as any)._oaEpoch === Date.UTC(1904, 0, 1), '1904 epoch set');
    await r.close();
}

async function testPhoneticAndSi(): Promise<void> {
    const sst = '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3"><si><t>plain</t></si><si><r><t>rich</t></r><rPh><t>phon</t></rPh></si><si/></sst>';
    const vals = parseSharedStringsXml(sst);
    assert(vals.length === 3, `sst len 3 got ${vals.length}`);
    assert(vals[0] === 'plain', 'plain');
    assert(vals[1] === 'rich', `phon skipped got ${vals[1]}`);
    assert(vals[2] === '', 'self-closing si empty');
}

async function testLineEndNormalization(): Promise<void> {
    // Raw CRLF / lone CR in element text are normalized to LF (XML 1.0 section 2.11).
    const sst = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>a\r\nb\rc</t></si><si><t>x&#13;y</t></si></sst>';
    const vals = parseSharedStringsXml(sst);
    assert(vals[0] === 'a\nb\nc', `CRLF/CR normalized got ${JSON.stringify(vals[0])}`);
    assert(vals[1] === 'x\ry', `char reference keeps CR got ${JSON.stringify(vals[1])}`);
    assert(unescapeXmlText('p\r\nq') === 'p\nq', 'unescapeXmlText normalizes raw CRLF');
    assert(escapeXmlText('p\r\nq') === 'p&#13;\nq', 'writer escapes CR as char reference');
    assert(unescapeXmlText(escapeXmlText('m\rn')) === 'm\rn', 'CR round-trips through writer and reader');
}

// Excel writes CR inside shared strings as _x000D_ (ST_Xstring); a literal _xHHHH_
// typed by a user must be written escaped so that it reads back unchanged.
async function testXstringEscapes(): Promise<void> {
    const excelSst = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>line1_x000D_\nline2</t></si><si><t>_x005F_x0041_</t></si></sst>';
    const excelVals = parseSharedStringsXml(excelSst);
    assert(excelVals[0] === 'line1\r\nline2', `_x000D_ decoded to CR got ${JSON.stringify(excelVals[0])}`);
    assert(excelVals[1] === '_x0041_', `_x005F_ decodes to a literal underscore got ${JSON.stringify(excelVals[1])}`);

    for (const literal of ['lit _x0041_ end', 'crlf\r\nand _x000D_', '_x005F_']) {
        const sst = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>${escapeXlsxText(literal)}</t></si></sst>`;
        assert(parseSharedStringsXml(sst)[0] === literal, `escapeXlsxText round-trips ${JSON.stringify(literal)}`);
    }

    const f = path.join(outputDir, 'xstring-escapes.xlsx');
    const writer = new XlsxWriter(f);
    writer.addSheet('Sheet1');
    writer.writeSheet([['line1\r\nline2', '_x0041_']], ['text', 'literal'], false);
    await writer.finalize();
    const reader = new XlsxReader();
    await reader.open(f);
    await reader.selectSheet(reader.getSheetNames()[0]);
    assert(await reader.read(), 'xstring file has a header row');
    assert(await reader.read(), 'xstring file has a data row');
    assert(reader.getValue(0) === 'line1\r\nline2', `file CRLF round-trips got ${JSON.stringify(reader.getValue(0))}`);
    assert(reader.getValue(1) === '_x0041_', `file literal _x0041_ round-trips got ${JSON.stringify(reader.getValue(1))}`);
    reader.close();
}

async function testCellTypes(): Promise<void> {
    const f = path.join(outputDir, 'celltypes.xlsx');
    createMinimalXlsx(f, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="str"><v>hello</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1" t="d"><v>2023-01-15T00:00:00</v></c><c r="D1" t="inlineStr"><is><r><t>he</t></r><r><t>llo</t></r></is></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
    });
    const r = new XlsxReader();
    await r.open(f);
    assert(await r.read(), 'types row');
    assert(r.getValue(0) === 'hello', `str type got ${r.getValue(0)}`);
    assert(r.getValue(1) === '#N/A', `e type got ${r.getValue(1)}`);
    const d = r.getValue(2);
    assert(d instanceof Date, `d type Date got ${d}`);
    // No timezone in the value: must be interpreted as UTC, like OA serials.
    assert((d as Date).getTime() === Date.UTC(2023, 0, 15), `d type UTC got ${(d as Date).toISOString()}`);
    assert(r.getValue(3) === 'hello', `inlineStr multi-t got ${r.getValue(3)}`);
    await r.close();
}

async function testSelfClosingRow(): Promise<void> {
    const f = path.join(outputDir, 'selfclose.xlsx');
    createMinimalXlsx(f, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"/><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
    });
    const r = new XlsxReader();
    await r.open(f);
    const rows: unknown[][] = [];
    while (await r.read()) rows.push([r.getValue(0)]);
    assert(rows.length === 3, `self-closing yields 3 rows got ${rows.length}`);
    assert(rows[0][0] === 1 && rows[1][0] == null && rows[2][0] === 3, `self-closing values ${JSON.stringify(rows)}`);
    await r.close();
}

async function testRowBreaksIsNotARow(): Promise<void> {
    const f = path.join(outputDir, 'rowbreaks.xlsx');
    createMinimalXlsx(f, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData><rowBreaks count="0" manualBreakCount="0"/></worksheet>',
            rId: 'rId1',
        }],
    });
    const r = new XlsxReader();
    await r.open(f);
    const rows: unknown[][] = [];
    while (await r.read()) rows.push([r.getValue(0)]);
    assert(rows.length === 2, `<rowBreaks/> must not add a row, got ${rows.length}`);
    assert(rows[0][0] === 1 && rows[1][0] === 2, `rowBreaks values ${JSON.stringify(rows)}`);
    await r.close();
}

async function testEmptyRowsGap(): Promise<void> {
    const f = path.join(outputDir, 'gaps.xlsx');
    createMinimalXlsx(f, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c></row><row r="5"><c r="A5"><v>5</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
    });
    const r = new XlsxReader();
    await r.open(f);
    const rows: unknown[][] = [];
    while (await r.read()) rows.push([r.getValue(0)]);
    assert(rows.length === 5, `gaps yield 5 rows got ${rows.length} ${JSON.stringify(rows)}`);
    assert(rows[0][0] === 1 && rows[1][0] == null && rows[2][0] === 3 && rows[3][0] == null && rows[4][0] === 5, 'gap values');
    await r.close();

    // Leading gaps.
    const f2 = path.join(outputDir, 'leading.xlsx');
    createMinimalXlsx(f2, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
    });
    const r2 = new XlsxReader();
    await r2.open(f2);
    const rows2: unknown[][] = [];
    while (await r2.read()) rows2.push([r2.getValue(0)]);
    assert(rows2.length === 3, `leading gaps 3 rows got ${rows2.length}`);
    assert(rows2[0][0] == null && rows2[1][0] == null && rows2[2][0] === 3, 'leading values');
    await r2.close();
}

async function testSelectSheet(): Promise<void> {
    const f = path.join(outputDir, 'multi.xlsx');
    const w = new XlsxWriter(f);
    w.addSheet('first');
    w.writeSheet([[1]], ['H'], false);
    w.addSheet('second');
    w.writeSheet([[2]], ['H2'], false);
    await w.finalize();

    const r = new XlsxReader();
    await r.open(f);
    assert(r.getSheetNames().length === 2, '2 sheets');
    // Default sheet 0.
    assert(await r.read(), 'read first sheet header');
    assert(r.actualSheetName === 'first', `default sheet first got ${r.actualSheetName}`);
    await r.selectSheet('second');
    assert(r.actualSheetName === 'second', 'select by name');
    assert(r.fieldCount === 0, 'fieldCount reset');
    assert(await r.read(), 'read second');
    assert(r.getValue(0) === 'H2', `second header ${r.getValue(0)}`);
    await r.selectSheet(0);
    assert(await r.read(), 're-select first');
    assert(r.actualSheetName === 'first', 'back to first');
    // nextResult
    await r.selectSheet(0);
    // consume first sheet fully
    while (await r.read()) { /* drain */ }
    // nextResult should go to second
    const hasNext = await r.nextResult();
    assert(hasNext, 'nextResult true');
    assert(r.actualSheetName === 'second', 'nextResult sheet');
    const hasNext2 = await r.nextResult();
    assert(!hasNext2, 'nextResult false at end');
    await r.close();

    // XLSB selectSheet
    const fb = path.join(outputDir, 'multi.xlsb');
    const wb = new XlsbWriter(fb);
    wb.addSheet('first');
    wb.writeSheet([[1]], ['H'], false);
    wb.addSheet('second');
    wb.writeSheet([[2]], ['H2'], false);
    await wb.finalize();
    const rb = new XlsbReader();
    await rb.open(fb);
    await rb.selectSheet(1);
    assert(rb.actualSheetName === 'second', 'xlsb select');
    assert(await rb.read(), 'xlsb read second');
    await rb.close();

    // Selecting a valid but empty worksheet must not throw.
    const fe = path.join(outputDir, 'empty-sheet.xlsb');
    const we = new XlsbWriter(fe);
    we.startSheet('empty', 1);
    we.endSheet();
    we.startSheet('data', 1);
    we.writeRow([7]);
    we.endSheet();
    await we.finalize();
    const re = new XlsbReader();
    await re.open(fe);
    await re.selectSheet('empty');
    assert(re.actualSheetName === 'empty', 'xlsb empty sheet selected');
    assert(!(await re.read()), 'xlsb empty sheet has no rows');
    assert(await re.nextResult(), 'xlsb nextResult after empty sheet');
    assert(await re.read(), 'xlsb reads sheet after empty one');
    assert(re.getValue(0) === 7, `xlsb value after empty sheet got ${re.getValue(0)}`);
    await re.close();
}

async function main(): Promise<void> {
    await testUnescape();
    await testParseTagAttributes();
    await testAbsolutePathAndAttrOrder();
    await testPhoneticAndSi();
    await testLineEndNormalization();
    await testXstringEscapes();
    await testCellTypes();
    await testSelfClosingRow();
    await testRowBreaksIsNotARow();
    await testEmptyRowsGap();
    await testSelectSheet();
    console.log('Reader XML fixes tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
