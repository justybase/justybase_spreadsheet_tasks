import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import { parseSharedStringsBin, readRecord } from '../../src/biff12Utils';
import { mkTempDir } from '../helpers/minimalXlsx';

const outputDir = mkTempDir('spreadsheet-sst-dim-');

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

async function testSstOrder(): Promise<void> {
    const f = path.join(outputDir, 'sst.xlsb');
    const w = new XlsbWriter(f);
    w.addSheet('data1');
    w.writeSheet([['a', 'a', 'b']], null, false);
    await w.finalize();

    const zip = new AdmZip(f);
    const sstBuf = zip.getEntry('xl/sharedStrings.bin')!.getData();
    const parsed = parseSharedStringsBin(sstBuf);
    // 3 total refs (a,a,b), 2 unique (a,b).
    assert(parsed.total === 3, `SST total 3 got ${parsed.total}`);
    assert(parsed.unique === 2, `SST unique 2 got ${parsed.unique}`);
    assert(parsed.values.includes('a') && parsed.values.includes('b'), 'SST values');
}

async function testDimPatch(): Promise<void> {
    const f = path.join(outputDir, 'dim.xlsb');
    const w = new XlsbWriter(f);
    w.addSheet('data1');
    w.writeSheet([[1, 2], [3, 4]], ['A', 'B'], false);
    await w.finalize();

    const out = path.join(outputDir, 'dim-updated.xlsb');
    const updater = new XlsbUpdater(f);
    updater.replaceSheetData('data1', [[9, 8, 7], [6, 5, 4], [3, 2, 1], [0, 0, 0]], { headers: ['X', 'Y', 'Z'] });
    updater.save(out);

    const zip = new AdmZip(out);
    const sheetEntry = zip.getEntry('xl/worksheets/sheet1.bin');
    assert(!!sheetEntry, 'sheet exists');
    const buf = sheetEntry!.getData();
    const rec: any = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    let dimFound = false;
    let selPatched = false;
    while (readRecord(buf, pos, rec)) {
        if (rec.id === 0x94) {
            dimFound = true;
            const rwLast = buf.readInt32LE(rec.dataStart + 4);
            const colLast = buf.readInt32LE(rec.dataStart + 12);
            // 4 data rows + header = 5 rows -> last index 4; 3 cols -> last 2.
            assert(rwLast === 4, `dim rwLast 4 got ${rwLast}`);
            assert(colLast === 2, `dim colLast 2 got ${colLast}`);
        }
        if (rec.id === 0x98) {
            // BrtSel must NOT be patched as dimension (old bug wrote rwLast/colLast there).
            // Its payload should not equal dim values at offsets 24/32 after fix.
            // We just ensure we did not write dim values into Sel: check that Sel at 24/32 is not 4/2
            // unless the original template already had those (it has 0/0).
            const v24 = rec.len >= 36 ? buf.readInt32LE(rec.dataStart + 24) : null;
            const v32 = rec.len >= 36 ? buf.readInt32LE(rec.dataStart + 32) : null;
            if (v24 === 4 && v32 === 2) {
                selPatched = true;
            }
        }
        pos = rec.dataEnd;
        if (pos >= buf.length) break;
    }
    assert(dimFound, 'BrtWsDim 0x94 found and patched');
    assert(!selPatched, 'BrtSel 0x98 must not be patched as dimension');
}

async function main(): Promise<void> {
    await testSstOrder();
    await testDimPatch();
    console.log('XLSB SST/dim tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
