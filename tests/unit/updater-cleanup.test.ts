import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxUpdater } from '../../src/XlsxUpdater';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { writeAdmZipStreaming } from '../../src/streamingZip';
import { createMinimalXlsx, zipHas, zipText, mkTempDir } from '../helpers/minimalXlsx';

const outputDir = mkTempDir('spreadsheet-cleanup-');

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

async function testCalcChainRemovedXlsx(): Promise<void> {
    const src = path.join(outputDir, 'calc.xlsx');
    createMinimalXlsx(src, {
        extraFiles: {
            'xl/calcChain.xml': '<?xml version="1.0"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1"/></calcChain>',
            'xl/workbook.xml': '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>',
            '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/><Override PartName="/xl/workbook.xml" ContentType="x"/></Types>',
            'xl/_rels/workbook.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="/xl/calcChain.xml"/></Relationships>',
        },
    });
    const out = path.join(outputDir, 'calc-out.xlsx');
    const updater = new XlsxUpdater(src);
    updater.replaceSheetData('data1', [[1]], { headers: ['H'] });
    updater.save(out);
    assert(!zipHas(out, 'xl/calcChain.xml'), 'calcChain removed');
    const ct = zipText(out, '[Content_Types].xml');
    assert(!ct.includes('calcChain'), 'Content_Types calcChain removed');
    const rels = zipText(out, 'xl/_rels/workbook.xml.rels');
    assert(!rels.includes('calcChain'), 'rels calcChain removed (absolute Target handled)');
    const wb = zipText(out, 'xl/workbook.xml');
    assert(wb.includes('fullCalcOnLoad'), 'calcPr requests recalculation');
}

async function testExistingCalcPrNotCorrupted(): Promise<void> {
    // Typical Excel workbook: <calcPr> exists without fullCalcOnLoad, and there
    // is no calcChain. Replacing data must not truncate the workbook XML.
    const src = path.join(outputDir, 'calcpr.xlsx');
    createMinimalXlsx(src, {
        workbookXml: '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>',
    });
    const out = path.join(outputDir, 'calcpr-out.xlsx');
    const updater = new XlsxUpdater(src);
    updater.replaceSheetData('data1', [[1]], { headers: ['H'] });
    updater.save(out);
    const wb = zipText(out, 'xl/workbook.xml');
    assert(wb.trimEnd().endsWith('</workbook>'), `workbook XML must stay complete, got ${wb}`);
    assert(wb.includes('fullCalcOnLoad="1"'), 'calcPr requests recalculation');
}

async function testTableRefUpdated(): Promise<void> {
    const src = path.join(outputDir, 'table.xlsx');
    createMinimalXlsx(src, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
        sharedStringsXml: '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>A</t></si><si><t>B</t></si></sst>',
        extraFiles: {
            'xl/tables/table1.xml': '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2"><autoFilter ref=\'A1:B2\'/><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>',
            'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        },
    });
    const out = path.join(outputDir, 'table-out.xlsx');
    const updater = new XlsxUpdater(src);
    updater.replaceSheetData('data1', [[1, 2], [3, 4], [5, 6]], { headers: ['A', 'B'] });
    updater.save(out);
    const tableXml = zipText(out, 'xl/tables/table1.xml');
    assert(tableXml.includes('ref="A1:B4"'), `table ref updated to A1:B4 got ${tableXml.substring(0, 300)}`);
    assert(/<autoFilter ref=['"]A1:B4['"]\/>/.test(tableXml), `single-quoted autoFilter updated got ${tableXml.substring(0, 300)}`);

    // Non-A1 tables must be left untouched (moving them corrupts the sheet).
    const src2 = path.join(outputDir, 'table-off.xlsx');
    createMinimalXlsx(src2, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="C1:D2"/><sheetData><row r="1"><c r="C1"><v>10</v></c><c r="D1"><v>20</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
        extraFiles: {
            'xl/tables/table1.xml': '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="C1:D2"><autoFilter ref="C1:D2"/></table>',
            'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        },
    });
    const out2 = path.join(outputDir, 'table-off-out.xlsx');
    const updater2 = new XlsxUpdater(src2);
    updater2.replaceSheetData('data1', [[1, 2]], { headers: ['a', 'b'] });
    updater2.save(out2);
    const offXml = zipText(out2, 'xl/tables/table1.xml');
    assert(offXml.includes('ref="C1:D2"'), `non-A1 table untouched got ${offXml.substring(0, 200)}`);

    // A width change would require rebuilding tableColumns, so leave the table
    // untouched instead of writing an inconsistent ref/count pair.
    const src3 = path.join(outputDir, 'table-width.xlsx');
    createMinimalXlsx(src3, {
        sheets: [{
            name: 'data1',
            xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row></sheetData></worksheet>',
            rId: 'rId1',
        }],
        extraFiles: {
            'xl/tables/table1.xml': '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2"><autoFilter ref="A1:B2"/><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>',
            'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        },
    });
    const out3 = path.join(outputDir, 'table-width-out.xlsx');
    const updater3 = new XlsxUpdater(src3);
    updater3.replaceSheetData('data1', [[1, 2, 3]], { headers: ['A', 'B', 'C'] });
    updater3.save(out3);
    const widthXml = zipText(out3, 'xl/tables/table1.xml');
    assert(widthXml.includes('ref="A1:B2"'), `width-changing table left untouched got ${widthXml.substring(0, 300)}`);
    assert(widthXml.includes('count="2"'), 'tableColumns count unchanged');
}

async function testCalcChainRemovedXlsb(): Promise<void> {
    const src = path.join(outputDir, 'calc.xlsb');
    const w = new XlsbWriter(src);
    w.addSheet('data1');
    w.writeSheet([[1]], ['H'], false);
    await w.finalize();
    // Inject fake calcChain.
    const zip = new AdmZip(src);
    zip.addFile('xl/calcChain.bin', Buffer.from([0x01, 0x02]));
    const relsPath = 'xl/_rels/workbook.bin.rels';
    const relsEntry = zip.getEntry(relsPath);
    let rels = relsEntry ? relsEntry.getData().toString('utf8') : '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    rels = rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.bin"/></Relationships>');
    zip.updateFile(relsPath, Buffer.from(rels, 'utf8'));
    const ctPath = '[Content_Types].xml';
    const ct = zip.getEntry(ctPath)!.getData().toString('utf8');
    zip.updateFile(ctPath, Buffer.from(ct.replace('</Types>', '<Override PartName="/xl/calcChain.bin" ContentType="application/vnd.ms-excel.calcChain"/></Types>'), 'utf8'));
    const src2 = path.join(outputDir, 'calc2.xlsb');
    fs.writeFileSync(src2, zip.toBuffer());

    const out = path.join(outputDir, 'calc2-out.xlsb');
    const updater = new XlsbUpdater(src2);
    updater.replaceSheetData('data1', [[2]], { headers: ['H'] });
    updater.save(out);
    assert(!zipHas(out, 'xl/calcChain.bin'), 'xlsb calcChain removed');
    assert(!zipText(out, '[Content_Types].xml').includes('calcChain.bin'), 'xlsb Content_Types calcChain removed');
}

async function testStreamingErrorCleansTemp(): Promise<void> {
    const src = path.join(outputDir, 'stream.xlsx');
    createMinimalXlsx(src);
    const zip = new AdmZip(src);
    const staged = new Map<string, string>();
    // Point to non-existent file to force append error.
    staged.set('xl/worksheets/sheet1.xml', path.join(outputDir, 'does-not-exist-12345.tmp'));
    const out = path.join(outputDir, 'stream-out.xlsx');
    let threw = false;
    try {
        await writeAdmZipStreaming(zip, staged, out);
    } catch {
        threw = true;
    }
    assert(threw, 'streaming with missing staged file throws');
    assert(!fs.existsSync(out), 'failed streaming does not leave target');
    const leftovers = fs.readdirSync(outputDir).filter(n => n.endsWith('.tmp'));
    assert(leftovers.length === 0, `no tmp left, got ${leftovers}`);
}

async function testAtomicFinalize(): Promise<void> {
    const f = path.join(outputDir, 'atomic.xlsx');
    const w = new XlsxWriter(f);
    w.addSheet('s');
    w.writeSheet([[1, 2]], ['A', 'B'], false);
    await w.finalize();
    assert(fs.existsSync(f), 'atomic target exists');
    const leftovers = fs.readdirSync(outputDir).filter(n => n.includes('.tmp'));
    assert(leftovers.length === 0, `no tmp after finalize, got ${leftovers}`);
    // Second finalize must reject (already finalized).
    let threw = false;
    try {
        await w.finalize();
    } catch {
        threw = true;
    }
    assert(threw, 'second finalize rejects');

    if (process.platform !== 'win32') {
        // Overwriting an existing file must keep its permissions when the
        // temp file is renamed into place.
        const modePath = path.join(outputDir, 'mode.xlsx');
        const first = new XlsxWriter(modePath);
        first.addSheet('s');
        first.writeSheet([[1]]);
        await first.finalize();
        fs.chmodSync(modePath, 0o600);
        const second = new XlsxWriter(modePath);
        second.addSheet('s');
        second.writeSheet([[2]]);
        await second.finalize();
        assert((fs.statSync(modePath).mode & 0o777) === 0o600, 'file mode preserved on overwrite');
    }
}

async function main(): Promise<void> {
    await testCalcChainRemovedXlsx();
    await testExistingCalcPrNotCorrupted();
    await testTableRefUpdated();
    await testCalcChainRemovedXlsb();
    await testStreamingErrorCleansTemp();
    await testAtomicFinalize();
    console.log('Updater cleanup/atomic tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
