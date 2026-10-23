import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsbWriter } from '../../src/XlsbWriter';
import { BrowserXlsxWriter } from '../../src/browser/BrowserXlsxWriter';
import { BrowserXlsbWriter } from '../../src/browser/BrowserXlsbWriter';
import { columnIndexToLetter } from '../../src/xmlUtils';
import { MAX_COLS, MAX_ROWS, MAX_CELL_CHARS } from '../../src/writerHelpers';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-limits-'));

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

async function testColumnLetters(): Promise<void> {
    assert(columnIndexToLetter(701) === 'ZZ', `701->ZZ got ${columnIndexToLetter(701)}`);
    assert(columnIndexToLetter(702) === 'AAA', `702->AAA got ${columnIndexToLetter(702)}`);
    assert(columnIndexToLetter(16383) === 'XFD', `16383->XFD got ${columnIndexToLetter(16383)}`);

    const f = path.join(outputDir, 'col.xlsx');
    const w = new XlsxWriter(f);
    assert((w as any)._getColumnLetter(701) === 'ZZ', 'XlsxWriter 701');
    assert((w as any)._getColumnLetter(702) === 'AAA', 'XlsxWriter 702 (was A before fix)');
    assert((w as any)._getColumnLetter(16383) === 'XFD', 'XlsxWriter max');
    try { (w as any).archive.abort(); } catch { /* ignore */ }
    try { (w as any).output.destroy(); } catch { /* ignore */ }

    const b = new BrowserXlsxWriter();
    assert((b as any)._cl(702) === 'AAA', 'BrowserXlsxWriter 702');
}

async function testLimitsThrow(): Promise<void> {
    // Column count over limit.
    {
        const f = path.join(outputDir, 'overcols.xlsx');
        const w = new XlsxWriter(f);
        w.addSheet('s');
        let threw = false;
        try {
            w.startSheet('s2', MAX_COLS + 1);
        } catch {
            threw = true;
        }
        assert(threw, 'startSheet over MAX_COLS throws');
        try { (w as any).archive.abort(); } catch { /* ignore */ }
        try { (w as any).output.destroy(); } catch { /* ignore */ }
    }

    // Row/header longer than columnCount.
    {
        const f = path.join(outputDir, 'rowfit.xlsx');
        const w = new XlsxWriter(f);
        w.addSheet('s');
        let threw = false;
        try {
            w.writeSheet([[1, 2, 3]], ['only-two', 'headers', 'here', 'extra']);
        } catch {
            threw = true;
        }
        // writeSheet derives columnCount from rows[0] (3) but headers has 4 -> must throw.
        assert(threw, 'writeSheet header longer than columnCount throws');
        try { (w as any).archive.abort(); } catch { /* ignore */ }
        try { (w as any).output.destroy(); } catch { /* ignore */ }
    }

    // Cell text over limit (UTF-16 units; astral counts as 2).
    {
        const longText = 'a'.repeat(MAX_CELL_CHARS + 1);
        const f = path.join(outputDir, 'longtext.xlsx');
        const w = new XlsxWriter(f);
        w.addSheet('s');
        let threw = false;
        try {
            w.writeSheet([[longText]]);
        } catch (e) {
            threw = String(e).includes('32767');
        }
        assert(threw, 'writeSheet overlong cell throws');
        try { (w as any).archive.abort(); } catch { /* ignore */ }
        try { (w as any).output.destroy(); } catch { /* ignore */ }

        const emoji = '😀'.repeat(16384); // 32768 UTF-16 units
        assert(emoji.length === 32768, 'emoji length is UTF-16 units');
        const f2 = path.join(outputDir, 'emoji.xlsb');
        const wb = new XlsbWriter(f2);
        wb.addSheet('s');
        let threw2 = false;
        try {
            wb.writeSheet([[emoji]]);
        } catch {
            threw2 = true;
        }
        assert(threw2, 'XLSB overlong astral cell throws');
        try { (wb as any).archive.abort(); } catch { /* ignore */ }
        try { (wb as any).output.destroy(); } catch { /* ignore */ }
    }

    // Browser writers enforce the same.
    {
        const b = new BrowserXlsxWriter();
        b.addSheet('s');
        let threw = false;
        try {
            b.writeSheet([['x'.repeat(MAX_CELL_CHARS + 1)]]);
        } catch {
            threw = true;
        }
        assert(threw, 'BrowserXlsxWriter overlong throws');

        const bb = new BrowserXlsbWriter();
        bb.addSheet('s');
        let threwB = false;
        try {
            (bb as any).startSheet('t', MAX_COLS + 1, undefined);
        } catch {
            threwB = true;
        }
        assert(threwB, 'BrowserXlsbWriter over MAX_COLS throws');
    }

    // MAX_ROWS constant sanity.
    assert(MAX_ROWS === 1_048_576, 'MAX_ROWS');
    assert(MAX_COLS === 16_384, 'MAX_COLS');
    assert(MAX_CELL_CHARS === 32_767, 'MAX_CELL_CHARS');
}

async function testWrite702Cols(): Promise<void> {
    const f = path.join(outputDir, 'wide.xlsx');
    const w = new XlsxWriter(f);
    w.addSheet('wide');
    const headers = Array.from({ length: 703 }, (_, i) => `H${i}`);
    const row = Array.from({ length: 703 }, (_, i) => i);
    w.writeSheet([row], headers, false);
    await w.finalize();
    assert(fs.existsSync(f), 'wide file exists');
    assert(!fs.existsSync(f + '.tmp'), 'no temp left');
    // Check dimension contains AAA (col 703).
    const AdmZip = (await import('adm-zip')).default;
    const zip = new AdmZip(f);
    const sheetXml = zip.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
    assert(sheetXml.includes('AAA'), `wide sheet uses AAA, got ${sheetXml.substring(0, 500)}`);
}

async function main(): Promise<void> {
    await testColumnLetters();
    await testLimitsThrow();
    await testWrite702Cols();
    console.log('Writer limits/columns tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
