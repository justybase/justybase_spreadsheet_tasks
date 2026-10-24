import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsxWriter } from '../../src/XlsxWriter';
import { XlsxReader } from '../../src/XlsxReader';
import { XlsxUpdater } from '../../src/XlsxUpdater';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-xlsx-positions-'));

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

function sheetXml(filePath: string): string {
    return new AdmZip(filePath).readAsText('xl/worksheets/sheet1.xml');
}

async function readRows(filePath: string): Promise<unknown[][]> {
    const reader = new XlsxReader();
    await reader.open(filePath);
    await reader.selectSheet(reader.getSheetNames()[0]);
    const rows: unknown[][] = [];
    while (await reader.read()) {
        const row: unknown[] = [];
        for (let i = 0; i < reader.fieldCount; i++) row.push(reader.getValue(i));
        rows.push(row);
    }
    reader.close();
    return rows;
}

// Cells that directly follow the previous cell carry no r attribute. A row
// starting outside column A and any cell after a gap keep r.
async function testOmitsRefsOnlyWithoutGaps(): Promise<void> {
    const file = path.join(outputDir, 'positions.xlsx');
    const writer = new XlsxWriter(file);
    writer.addSheet('Sheet1');
    writer.writeSheet([[1, 'x', null, 4], [null, 'y', 3, 4]], ['h1', 'h2', 'h3', 'h4'], false);
    await writer.finalize();

    const xml = sheetXml(file);
    // Row 1 is the header, row 2 = [1, 'x', null, 4], row 3 = [null, 'y', 3, 4].
    assert(xml.includes('<row r="1"><c t="s" s="3"><v>0</v></c>'), 'header A1 omits r');
    assert(!xml.includes('r="A1"') && !xml.includes('r="B1"') && !xml.includes('r="C1"'),
        'header cells are contiguous from A and omit r');
    assert(xml.includes('<row r="2"><c><v>1</v></c><c t="s"><v>4</v></c><c r="D2"><v>4</v></c></row>'),
        'row 2: A2 and B2 omit r, D2 keeps r after the null gap in C2');
    assert(xml.includes('<row r="3"><c r="B3" t="s"><v>5</v></c><c><v>3</v></c><c><v>4</v></c></row>'),
        'row 3: B3 keeps r because the row starts with a null in A3; the cells after it omit r');

    const rows = await readRows(file);
    assert(JSON.stringify(rows) === JSON.stringify([
        ['h1', 'h2', 'h3', 'h4'],
        [1, 'x', null, 4],
        [null, 'y', 3, 4],
    ]), `values and positions round-trip, got ${JSON.stringify(rows)}`);
    console.log('  xlsx writeSheet: refs omitted only where contiguous, values round-trip');
}

async function testStreamingRefs(): Promise<void> {
    const file = path.join(outputDir, 'stream.xlsx');
    const writer = new XlsxWriter(file);
    writer.startSheet('Data', 3, ['a', 'b', 'c'], { doAutofilter: false });
    writer.writeRow([null, 'y', 3]);
    writer.writeRow([1, null, 2]);
    writer.endSheet();
    await writer.finalize();

    const xml = sheetXml(file);
    assert(xml.includes('<row r="2"><c r="B2" t="s">'), 'streaming: B2 keeps r after null A2');
    assert(!xml.includes('r="C2"'), 'streaming: C2 omits r');
    assert(xml.includes('<row r="3"><c><v>1</v></c><c r="C3"><v>2</v></c></row>'),
        'streaming: gap after A3 is marked with r on C3');

    const rows = await readRows(file);
    assert(JSON.stringify(rows) === JSON.stringify([['a', 'b', 'c'], [null, 'y', 3], [1, null, 2]]),
        `streaming round-trip, got ${JSON.stringify(rows)}`);
    console.log('  xlsx streaming writer: refs and round-trip');
}

// An updater that inherits styles must still see the data column of a file that
// omits r. Column B is a date column; the new value must inherit its style.
async function testUpdaterInheritsStylesWithoutRefs(): Promise<void> {
    const source = path.join(outputDir, 'styled-source.xlsx');
    const writer = new XlsxWriter(source);
    writer.addSheet('Sheet1');
    writer.writeSheet([[1, new Date(Date.UTC(2024, 0, 5))], [2, new Date(Date.UTC(2024, 0, 6))]], ['id', 'day'], false);
    await writer.finalize();
    assert(!sheetXml(source).includes('r="A2"'), 'source has no r on column A (precondition)');
    const sourceStyle = /<c s="(\d+)"><v>\d+<\/v><\/c>/.exec(sheetXml(source));
    assert(sourceStyle !== null && Number(sourceStyle[1]) > 0, 'source date cell has a style (precondition)');

    const updated = path.join(outputDir, 'styled-updated.xlsx');
    const updater = new XlsxUpdater(source);
    await updater.replaceSheetDataStream('Sheet1', [[5, new Date(Date.UTC(2024, 1, 1))]], { headers: ['id', 'day'] });
    updater.save(updated);
    const xml = sheetXml(updated);
    assert(/<c r="B2" s="\d+"><v>/.test(xml), `new date cell inherits its column style, got ${xml.slice(0, 400)}`);
    console.log('  xlsx updater: style inherited from a file whose cells omit r');
}

// Guard against the size regression: a full 10,000 x 20 block must not repeat the cell
// reference on every cell. A cell with a numeric value is `<c><v>NNNNNN</v></c>` (about 20 bytes).
async function testWideBlockSize(): Promise<void> {
    const rows: number[][] = [];
    for (let r = 0; r < 10000; r++) rows.push(Array.from({ length: 20 }, (_, c) => r * 20 + c));

    const file = path.join(outputDir, 'wide-10k.xlsx');
    const writer = new XlsxWriter(file);
    writer.addSheet('Sheet1');
    writer.writeSheet(rows, null, false);
    await writer.finalize();

    const xml = sheetXml(file);
    assert(!/<c r=/.test(xml), 'contiguous 10,000 x 20 block carries no cell references');
    const bytesPerCell = xml.length / (10000 * 20);
    assert(bytesPerCell < 24, `sheet XML stays near 20 bytes per cell, got ${bytesPerCell.toFixed(1)}`);
    console.log(`  xlsx 10,000 x 20 block: ${bytesPerCell.toFixed(1)} bytes of sheet XML per cell, no cell refs`);
}

async function testCompressionLevelOption(): Promise<void> {
    const rows: (string | number)[][] = [];
    for (let i = 0; i < 4000; i++) rows.push([i, `customer-${i % 97}`, i * 0.5]);

    const fast = path.join(outputDir, 'level1.xlsx');
    const writerFast = new XlsxWriter(fast, { compressionLevel: 1 });
    writerFast.addSheet('Sheet1');
    writerFast.writeSheet(rows, ['id', 'name', 'value'], false);
    await writerFast.finalize();

    const normal = path.join(outputDir, 'level6.xlsx');
    const writerNormal = new XlsxWriter(normal);
    writerNormal.addSheet('Sheet1');
    writerNormal.writeSheet(rows, ['id', 'name', 'value'], false);
    await writerNormal.finalize();

    assert(fs.statSync(normal).size < fs.statSync(fast).size,
        `default level 6 is smaller than level 1 (${fs.statSync(normal).size} vs ${fs.statSync(fast).size})`);
    assert(JSON.stringify(await readRows(normal)) === JSON.stringify(await readRows(fast)),
        'both compression levels round-trip the same values');

    let bad = false;
    try { new XlsxWriter(path.join(outputDir, 'bad.xlsx'), { compressionLevel: 10 }); } catch (e) { bad = e instanceof RangeError; }
    assert(bad, 'compressionLevel above 9 throws RangeError');
    console.log('  xlsx compressionLevel: level 6 smaller than level 1, same values, out-of-range rejected');
}

async function main(): Promise<void> {
    await testOmitsRefsOnlyWithoutGaps();
    await testStreamingRefs();
    await testUpdaterInheritsStylesWithoutRefs();
    await testWideBlockSize();
    await testCompressionLevelOption();
    console.log('XLSX cell position tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
