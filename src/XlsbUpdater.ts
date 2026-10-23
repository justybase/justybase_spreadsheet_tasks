import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { CellValue, getFormat, unwrapCell } from './Formats';
import { writeBufferAtomically } from './atomicFile';
import { writeAdmZipStreaming, StagedZipPart } from './streamingZip';
import {
    readRecord,
    buildRecord,
    readUtf16,
    parseSharedStringsBin,
    Biff12Record,
} from './biff12Utils';
import { BiffReaderWriter } from './BiffReaderWriter';
import { trimTrailingEmptyRows, parseTagAttributes, resolveRelTarget } from './xmlUtils';
import { timeToOaDate } from './writerHelpers';

/**
 * Options for {@link XlsbUpdater.replaceSheetData}.
 */
export interface ReplaceSheetDataOptions {
    /**
     * Optional header row written into row 1 of the target sheet.
     * When omitted, row 1 is filled with the data rows.
     */
    headers?: string[];
    /**
     * Style strategy for the new cells:
     * - `'inherit'` (default): reuse the dominant cell style (xf) of each column from the
     *   existing data (and the original header style for row 1), so date/number
     *   formatting of the sheet is preserved. Falls back to General (0).
     * - `'general'`: all new cells use xf 0 (General).
     */
    styleFallback?: 'inherit' | 'general';
}

/** A one-pass synchronous or asynchronous source of worksheet rows. */
export type RowSource =
    | Iterable<ReadonlyArray<CellValue>>
    | AsyncIterable<ReadonlyArray<CellValue>>;

/** Worksheet cell records (BrtCellBlank … BrtFmlaError). */
const CELL_RECORDS = new Set<number>([0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b]);

const RK_INT_LOWER = -1 << 29;
const RK_INT_UPPER = (1 << 29) - 1;

/**
 * Update the data of a worksheet inside an existing XLSB file without rebuilding
 * the workbook. Everything outside the target sheet's rows — pivot tables and
 * their caches, other sheets, styles, column widths, views — is preserved.
 * This makes it possible to refresh a data sheet that pivot tables are wired to.
 *
 * ```ts
 * const updater = new XlsbUpdater('report.xlsb');
 * updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
 * updater.save('report_new.xlsb'); // or updater.save() to overwrite in place
 * ```
 */
export class XlsbUpdater {
    private readonly zip: AdmZip;
    private readonly sourcePath: string;
    private temporaryDirectory: string | null = null;
    private readonly stagedParts: StagedZipPart = new Map();

    private sheetNameToPath: Map<string, string> = new Map();
    private pivotCacheDefPaths: string[] = [];

    private sharedStringsBuf: Buffer | null = null;
    private sharedStringsValues: string[] = [];
    private sharedStringsTotal: number = 0;
    private sharedStringsUnique: number = 0;
    private sharedStringsEndSst: number = 0;
    private stringIndexMap: Map<string, number> = new Map();
    private sharedStringsDirty: boolean = false;

    private _oaEpoch: number;

    /**
     * @param path Path to an existing `.xlsb` file.
     */
    constructor(filePath: string) {
        this.sourcePath = filePath;
        this._oaEpoch = Date.UTC(1899, 11, 30);

        if (!fs.existsSync(filePath)) {
            throw new Error(`XlsbUpdater: file not found: ${filePath}`);
        }

        this.zip = new AdmZip(filePath);

        if (!this.zip.getEntry('xl/workbook.bin')) {
            throw new Error(
                'XlsbUpdater: not an XLSB workbook (xl/workbook.bin missing). Convert the file to XLSB first.'
            );
        }

        this._loadWorkbookStructure();
    }

    /** Worksheet names in workbook order. */
    getSheetNames(): string[] {
        return [...this.sheetNameToPath.keys()];
    }

    /**
     * Clear the entire cell data of the target sheet and write new rows in its place.
     * All other parts of the workbook (pivot tables, other sheets, styles, …) are preserved.
     *
     * @param sheetName Name of the worksheet to replace.
     * @param rows New data rows (`CellValue` cells; null/undefined = blank).
     * @param options See {@link ReplaceSheetDataOptions}.
     */
    replaceSheetData(
        sheetName: string,
        rows: CellValue[][],
        options: ReplaceSheetDataOptions = {}
    ): void {
        const { headers, styleFallback = 'inherit' } = options;
        rows = trimTrailingEmptyRows(rows);

        const sheetPath = this.sheetNameToPath.get(sheetName);
        if (!sheetPath) {
            throw new Error(`XlsbUpdater: sheet "${sheetName}" not found in the workbook.`);
        }

        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsbUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const sheetBuf = this._entryData(sheetPath);
        this._ensureSharedStringsLoaded();
        if (this.sharedStringsBuf !== null) {
            this.sharedStringsTotal = Math.max(
                0,
                this.sharedStringsTotal - this._sharedStringRefCount(sheetBuf),
            );
            this.sharedStringsDirty = true;
        }

        let dataStyles: Map<number, number> = new Map();
        let headerStyles: Map<number, number> = new Map();
        let dateXf: number | null = null;

        if (styleFallback !== 'general') {
            const collected = this._collectStyles(sheetBuf);
            dataStyles = collected.dataStyles;
            headerStyles = collected.headerStyles;
            dateXf = this._findDateXf();
        }

        const lastCol = Math.max(0, Math.max(
            rows.reduce((m, r) => Math.max(m, r.length), 0),
            headers?.length ?? 0
        ) - 1);

        const newRows = this._buildRows(rows, headers, dataStyles, headerStyles, dateXf);
        const { rowsStart, rowsEnd, autoFilterRanges } = this._findRowsRegion(sheetBuf);

        const newSheet = Buffer.alloc(sheetBuf.length - (rowsEnd - rowsStart) + newRows.length);
        sheetBuf.copy(newSheet, 0, 0, rowsStart);
        newRows.copy(newSheet, rowsStart);
        sheetBuf.copy(newSheet, rowsStart + newRows.length, rowsEnd);

        const lastRowIdx = rows.length + (headers ? 1 : 0) - 1;
        const delta = newRows.length - (rowsEnd - rowsStart);
        for (const af of autoFilterRanges) {
            const newAf = af + delta;
            newSheet.writeInt32LE(lastCol, newAf + 12);
            newSheet.writeInt32LE(lastRowIdx, newAf + 4);
        }

        // Keep the sheet dimension record (0x94) in sync with the new data. Excel uses
        // its rwLast/colLast fields to read the full row range; stale values make the
        // last written row invisible (and drop it from pivot refreshes).
        this._patchDimension(newSheet, lastRowIdx, lastCol);

        this.zip.updateFile(sheetPath, newSheet);
        this._discardStagedPart(sheetPath);

        if (this.sharedStringsBuf !== null) {
            this._commitSharedStrings();
        }

        const rwLast = Math.max(0, rows.length + (headers ? 1 : 0) - 1);
        const colLast = lastCol;
        this._patchPivotCaches(sheetName, rwLast, colLast);
        this._removeCalcChain();
    }

    /** Replace worksheet rows from a one-pass source. */
    async replaceSheetDataStream(
        sheetName: string,
        rows: RowSource,
        options: ReplaceSheetDataOptions = {},
    ): Promise<void> {
        const { headers, styleFallback = 'inherit' } = options;
        if (styleFallback !== 'inherit' && styleFallback !== 'general') {
            throw new Error("XlsbUpdater: styleFallback must be 'inherit' or 'general'.");
        }

        const sheetPath = this.sheetNameToPath.get(sheetName);
        if (!sheetPath) {
            throw new Error(`XlsbUpdater: sheet "${sheetName}" not found in the workbook.`);
        }
        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsbUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const oldStage = this.stagedParts.get(sheetPath);
        const beforeValuesLength = this.sharedStringsValues.length;
        const beforeAppendCount = this._appendCount;
        const beforeSharedStringsTotal = this.sharedStringsTotal;
        const beforeSharedStringsUnique = this.sharedStringsUnique;
        const beforeSharedStringsEndSst = this.sharedStringsEndSst;
        const beforeSharedStringsDirty = this.sharedStringsDirty;
        const beforeStringIndexMap = new Map(this.stringIndexMap);
        const sharedStringsEntry = this.zip.getEntry('xl/sharedStrings.bin');
        const beforeSharedStringsBytes = sharedStringsEntry || this.stagedParts.has('xl/sharedStrings.bin')
            ? Buffer.from(this._entryData('xl/sharedStrings.bin'))
            : null;
        const beforeSharedStringsBuf = this.sharedStringsBuf ? Buffer.from(this.sharedStringsBuf) : null;
        const oldSheet = this._entryData(sheetPath);
        const rowsPath = this._temporaryPath('.rows');
        const outputPath = this._temporaryPath('.bin');
        const createdParts = [rowsPath, outputPath];
        let output: fs.WriteStream | null = null;

        try {
            this._ensureSharedStringsLoaded();
            if (this.sharedStringsBuf !== null) {
                this.sharedStringsTotal = Math.max(
                    0,
                    this.sharedStringsTotal - this._sharedStringRefCount(oldSheet),
                );
                this.sharedStringsDirty = true;
            }

            let dataStyles = new Map<number, number>();
            let headerStyles = new Map<number, number>();
            let dateXf: number | null = null;
            if (styleFallback !== 'general') {
                const collected = this._collectStyles(oldSheet);
                dataStyles = collected.dataStyles;
                headerStyles = collected.headerStyles;
                dateXf = this._findDateXf();
            }

            const rowOutput = fs.createWriteStream(rowsPath);
            output = rowOutput;
            let outputError: Error | null = null;
            rowOutput.on('error', error => { outputError = error; });
            let writtenBytes = 0;
            let lastKeptEnd = 0;
            let width = headers?.length ?? 0;
            let dataRowsSeen = 0;
            let keptDataRows = 0;
            let pendingEmptyWidth = 0;
            let nextRow = 0;
            const rowHeaderOffsets: number[] = [];

            const write = async (data: Buffer): Promise<void> => {
                if (outputError) throw outputError;
                writtenBytes += data.length;
                if (rowOutput.write(data)) return;
                await new Promise<void>((resolve, reject) => {
                    const onDrain = () => { cleanup(); resolve(); };
                    const onError = (error: Error) => { cleanup(); reject(error); };
                    const cleanup = () => {
                        rowOutput.off('drain', onDrain);
                        rowOutput.off('error', onError);
                    };
                    rowOutput.once('drain', onDrain);
                    rowOutput.once('error', onError);
                });
            };

            const writeRow = async (rowNumber: number, cells: Buffer[]): Promise<void> => {
                rowHeaderOffsets.push(writtenBytes);
                const header = Buffer.alloc(27);
                header[0] = 0x00;
                header[1] = 25;
                header.writeInt32LE(rowNumber, 2);
                header[10] = 0x2c;
                header[11] = 0x01;
                header[15] = 0x01;
                header.writeInt32LE(0, 19);
                header.writeInt32LE(0, 23);
                await write(header);
                for (const cell of cells) await write(cell);
            };

            if (headers !== undefined) {
                await writeRow(0, headers.map((value, column) =>
                    this._stringCellBytes(column, styleFallback === 'general' ? 0 : (headerStyles.get(column) ?? 0), value)));
                lastKeptEnd = writtenBytes;
                nextRow = 1;
            }

            for await (const sourceRow of rows) {
                if (typeof sourceRow === 'string' || sourceRow instanceof Uint8Array) {
                    throw new TypeError('Each row must be an iterable of cells, not text');
                }
                const row = Array.from(sourceRow) as CellValue[];
                const cells = row
                    .map((value, column) => value === null || value === undefined
                        ? null
                        : this._valueCellBytes(value, column, dataStyles, dateXf))
                    .filter((cell): cell is Buffer => cell !== null);
                await writeRow(nextRow, cells);
                dataRowsSeen++;
                nextRow++;

                if (row.every(value => value === null || value === undefined)) {
                    pendingEmptyWidth = Math.max(pendingEmptyWidth, row.length);
                } else {
                    width = Math.max(width, pendingEmptyWidth, row.length);
                    pendingEmptyWidth = 0;
                    keptDataRows = dataRowsSeen;
                    lastKeptEnd = writtenBytes;
                }
            }

            await new Promise<void>((resolve, reject) => {
                rowOutput.once('close', resolve);
                rowOutput.once('error', reject);
                rowOutput.end();
            });
            if (outputError) throw outputError;

            const rowsFd = fs.openSync(rowsPath, 'r+');
            try {
                fs.ftruncateSync(rowsFd, lastKeptEnd);
            } finally {
                fs.closeSync(rowsFd);
            }
            const lastCol = Math.max(0, width - 1);
            const rowPatchFd = fs.openSync(rowsPath, 'r+');
            try {
                const lastColBuffer = Buffer.alloc(4);
                lastColBuffer.writeInt32LE(lastCol, 0);
                for (const offset of rowHeaderOffsets) {
                    if (offset >= lastKeptEnd) continue;
                    fs.writeSync(rowPatchFd, lastColBuffer, 0, lastColBuffer.length, offset + 23);
                }
            } finally {
                fs.closeSync(rowPatchFd);
            }

            const { rowsStart, rowsEnd, autoFilterRanges } = this._findRowsRegion(oldSheet);
            const oldLength = oldSheet.length;
            const newRowsLength = fs.statSync(rowsPath).size;
            const delta = newRowsLength - (rowsEnd - rowsStart);
            const outputFd = fs.openSync(outputPath, 'w');
            try {
                fs.writeSync(outputFd, oldSheet, 0, rowsStart, null);
                const rowsFdRead = fs.openSync(rowsPath, 'r');
                try {
                    const chunk = Buffer.alloc(1024 * 1024);
                    let position = 0;
                    while (position < newRowsLength) {
                        const count = fs.readSync(rowsFdRead, chunk, 0, Math.min(chunk.length, newRowsLength - position), position);
                        if (count <= 0) throw new Error('XlsbUpdater: failed to read staged rows.');
                        fs.writeSync(outputFd, chunk, 0, count);
                        position += count;
                    }
                } finally {
                    fs.closeSync(rowsFdRead);
                }
                fs.writeSync(outputFd, oldSheet, rowsEnd, oldLength - rowsEnd, rowsStart + newRowsLength);
            } finally {
                fs.closeSync(outputFd);
            }

            this._patchSheetMetadataFile(
                outputPath,
                oldSheet,
                rowsEnd,
                delta,
                autoFilterRanges,
                Math.max(0, keptDataRows + (headers !== undefined ? 1 : 0) - 1),
                lastCol,
            );
            this._stagePart(sheetPath, outputPath);
            this._commitSharedStrings();
            this._patchPivotCaches(
                sheetName,
                Math.max(0, keptDataRows + (headers !== undefined ? 1 : 0) - 1),
                lastCol,
            );
            this._removeCalcChain();

            if (oldStage && oldStage !== outputPath) this._removeTemporaryFile(oldStage);
            createdParts.splice(createdParts.indexOf(outputPath), 1);
        } catch (error) {
            this.sharedStringsValues.length = beforeValuesLength;
            this._appendCount = beforeAppendCount;
            this.sharedStringsTotal = beforeSharedStringsTotal;
            this.sharedStringsUnique = beforeSharedStringsUnique;
            this.sharedStringsEndSst = beforeSharedStringsEndSst;
            this.sharedStringsDirty = beforeSharedStringsDirty;
            this.sharedStringsBuf = beforeSharedStringsBuf;
            this.stringIndexMap.clear();
            for (const [value, index] of beforeStringIndexMap) this.stringIndexMap.set(value, index);
            if (beforeSharedStringsBytes) {
                this.zip.updateFile('xl/sharedStrings.bin', beforeSharedStringsBytes);
            }
            if (oldStage) this.stagedParts.set(sheetPath, oldStage);
            else this.stagedParts.delete(sheetPath);
            throw error;
        } finally {
            if (output && !output.closed) output.destroy();
            for (const temporaryPath of createdParts) this._removeTemporaryFile(temporaryPath);
        }
    }

    /**
     * Write the updated workbook to disk.
     * @param outputPath Destination path; defaults to the source file (overwrites in place).
     */
    save(outputPath?: string): void {
        const target = outputPath ?? this.sourcePath;
        this._materializeStagedParts();
        writeBufferAtomically(this.zip.toBuffer(), target);
    }

    /** Save staged streaming replacements without materialising the ZIP. */
    async saveStreaming(outputPath?: string): Promise<void> {
        const target = outputPath ?? this.sourcePath;
        await writeAdmZipStreaming(this.zip, this.stagedParts, target);
        this._materializeStagedParts();
    }

    /** The updated workbook as an in-memory ZIP buffer. */
    toBuffer(): Buffer {
        this._materializeStagedParts();
        return this.zip.toBuffer();
    }

    /**
     * Release the staging directory and any unsaved staged parts.
     * The updater remains usable; a new staging directory is created lazily
     * on the next streaming replacement.
     */
    dispose(): void {
        for (const stagedPath of this.stagedParts.values()) {
            try {
                fs.rmSync(stagedPath, { force: true });
            } catch {
                // Best-effort cleanup.
            }
        }
        this.stagedParts.clear();
        if (this.temporaryDirectory !== null) {
            try {
                fs.rmSync(this.temporaryDirectory, { recursive: true, force: true });
            } catch {
                // Best-effort cleanup.
            }
            this.temporaryDirectory = null;
        }
    }

    /** Support `using` declarations for automatic staging cleanup. */
    [Symbol.dispose](): void {
        this.dispose();
    }

    private _ensureTemporaryDirectory(): string {
        if (this.temporaryDirectory === null) {
            this.temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-updater-'));
        } else {
            try {
                fs.mkdirSync(this.temporaryDirectory, { recursive: true });
            } catch {
                // Best-effort; file creation will surface persistent failures.
            }
        }
        return this.temporaryDirectory;
    }

    private _temporaryPath(suffix: string): string {
        return path.join(this._ensureTemporaryDirectory(), `part-${Date.now()}-${Math.random().toString(16).slice(2)}${suffix}`);
    }

    private _removeTemporaryFile(filePath: string): void {
        try {
            fs.rmSync(filePath);
        } catch (error) {
            if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
        }
    }

    private _entryData(name: string): Buffer {
        const stagedPath = this.stagedParts.get(name);
        if (stagedPath) return fs.readFileSync(stagedPath);
        const entry = this.zip.getEntry(name);
        if (!entry) throw new Error(`XlsbUpdater: ZIP member missing: ${name}`);
        return entry.getData();
    }

    private _stagePart(name: string, filePath: string): void {
        this.stagedParts.set(name, filePath);
    }

    private _discardStagedPart(name: string): void {
        const stagedPath = this.stagedParts.get(name);
        if (!stagedPath) return;
        this._removeTemporaryFile(stagedPath);
        this.stagedParts.delete(name);
    }

    private _materializeStagedParts(): void {
        for (const [name, stagedPath] of this.stagedParts) {
            this.zip.updateFile(name, fs.readFileSync(stagedPath));
            this._removeTemporaryFile(stagedPath);
        }
        this.stagedParts.clear();
    }

    // ------------------------------------------------------------------
    // Workbook structure
    // ------------------------------------------------------------------

    private _loadWorkbookStructure(): void {
        const relsEntry = this.zip.getEntry('xl/_rels/workbook.bin.rels');
        const rIdToTarget: Record<string, string> = {};
        if (relsEntry) {
            const rels = relsEntry.getData().toString('utf8');
            const relTagRe = /<Relationship\b[^>]*>/g;
            let tagMatch;
            while ((tagMatch = relTagRe.exec(rels)) !== null) {
                const attrs = parseTagAttributes(tagMatch[0]);
                const id = attrs['Id'] ?? attrs['id'];
                const target = attrs['Target'] ?? attrs['target'];
                if (id !== undefined && target !== undefined) {
                    rIdToTarget[id] = target;
                }
            }
        }

        const wbEntry = this.zip.getEntry('xl/workbook.bin');
        if (wbEntry) {
            const wbBuf = wbEntry.getData();
            const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
            let pos = 0;
            while (readRecord(wbBuf, pos, rec)) {
                pos = rec.dataEnd;
                if (rec.id !== 0x009c || rec.len < 8) continue;

                // BrtBundleSh: hidden(4) + sheetId(4) + rId(string) + name(string)
                let o = rec.dataStart + 8;
                const rIdLen = wbBuf.readUInt32LE(o);
                o += 4;
                const rId = readUtf16(wbBuf, o, rIdLen);
                o += rIdLen * 2;
                const nameLen = wbBuf.readUInt32LE(o);
                o += 4;
                const name = readUtf16(wbBuf, o, nameLen);

                let target = rIdToTarget[rId] ?? '';
                target = resolveRelTarget(target);
                if (target) {
                    this.sheetNameToPath.set(name, target);
                }
            }
            // 1904 date system: BrtWbProp (0x99) bit0 f1904.
            try {
                const rec1904: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
                let p1904 = 0;
                while (readRecord(wbBuf, p1904, rec1904)) {
                    if (rec1904.id === 0x99 && rec1904.len >= 4) {
                        const is1904 = (wbBuf[rec1904.dataStart] & 0x01) !== 0;
                        this._oaEpoch =
                            is1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
                        break;
                    }
                    p1904 = rec1904.dataEnd;
                }
            } catch {
                // Keep default epoch.
            }
        }

        for (const e of this.zip.getEntries()) {
            if (/^xl\/pivotCache\/pivotCacheDefinition\d*\.bin$/.test(e.entryName)) {
                this.pivotCacheDefPaths.push(e.entryName);
            }
        }
    }

    // ------------------------------------------------------------------
    // Shared strings
    // ------------------------------------------------------------------

    private _appendCount: number = 0;

    private _ensureSharedStringsLoaded(): void {
        if (this.stringIndexMap.size > 0 || this.sharedStringsBuf !== null) {
            return;
        }
        const entry = this.zip.getEntry('xl/sharedStrings.bin');
        if (!entry && !this.stagedParts.has('xl/sharedStrings.bin')) {
            this.sharedStringsBuf = null;
            return;
        }
        this.sharedStringsBuf = this._entryData('xl/sharedStrings.bin');
        const parsed = parseSharedStringsBin(this.sharedStringsBuf);
        this.sharedStringsValues = parsed.values;
        this.sharedStringsTotal = parsed.total;
        this.sharedStringsUnique = parsed.unique;
        this.sharedStringsEndSst = parsed.endSstOffset;

        this.sharedStringsValues.forEach((value, index) => {
            if (!this.stringIndexMap.has(value)) {
                this.stringIndexMap.set(value, index);
            }
        });
        this._appendCount = 0;
        this.sharedStringsDirty = false;
    }

    private _commitSharedStrings(): void {
        if (this.sharedStringsBuf === null || !this.sharedStringsDirty) return;

        const originalValuesLength = this.sharedStringsValues.length - this._appendCount;

        let newItems = Buffer.alloc(0);
        for (let i = originalValuesLength; i < this.sharedStringsValues.length; i++) {
            const text = this.sharedStringsValues[i];
            const payload = Buffer.alloc(1 + 4 + text.length * 2);
            payload[0] = 0x00;
            payload.writeUInt32LE(text.length, 1);
            payload.write(text, 5, text.length * 2, 'utf16le');
            newItems = Buffer.concat([newItems, buildRecord(0x0013, payload)]);
        }

        const out = Buffer.alloc(this.sharedStringsBuf.length + newItems.length);
        this.sharedStringsBuf.copy(out, 0, 0, this.sharedStringsEndSst);
        newItems.copy(out, this.sharedStringsEndSst);
        this.sharedStringsBuf.copy(out, this.sharedStringsEndSst + newItems.length, this.sharedStringsEndSst);

        // Patch cstTotal / cstUnique inside BrtBeginSst.
        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(out, pos, rec)) {
            if (rec.id === 0x009f) {
                out.writeUInt32LE(this.sharedStringsTotal, rec.dataStart);
                out.writeUInt32LE(this.sharedStringsUnique, rec.dataStart + 4);
                break;
            }
            pos = rec.dataEnd;
        }

        this.zip.updateFile('xl/sharedStrings.bin', out);
        this.sharedStringsBuf = out;
        this._appendCount = 0;
        this.sharedStringsDirty = false;
    }

    // ------------------------------------------------------------------
    // Styles
    // ------------------------------------------------------------------

    private _collectStyles(sheetBuf: Buffer): { dataStyles: Map<number, number>; headerStyles: Map<number, number> } {
        const dataCounts = new Map<number, Map<number, number>>();
        const headerStyles = new Map<number, number>();

        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        let row = -1;
        while (readRecord(sheetBuf, pos, rec)) {
            pos = rec.dataEnd;
            if (rec.id === 0x00) {
                row = sheetBuf.readInt32LE(rec.dataStart);
                continue;
            }
            if (!CELL_RECORDS.has(rec.id) || rec.len < 8) continue;
            const col = sheetBuf.readUInt32LE(rec.dataStart);
            const xf = sheetBuf.readUInt32LE(rec.dataStart + 4) & 0xffffff;
            if (xf <= 0) continue;

            if (row === 0) {
                if (!headerStyles.has(col)) headerStyles.set(col, xf);
                continue;
            }
            let colCounts = dataCounts.get(col);
            if (!colCounts) {
                colCounts = new Map();
                dataCounts.set(col, colCounts);
            }
            colCounts.set(xf, (colCounts.get(xf) ?? 0) + 1);
        }

        const dataStyles = new Map<number, number>();
        for (const [col, colCounts] of dataCounts) {
            let best = 0;
            let bestCount = -1;
            for (const [xf, count] of colCounts) {
                if (count > bestCount) {
                    best = xf;
                    bestCount = count;
                }
            }
            if (bestCount > 0) dataStyles.set(col, best);
        }

        return { dataStyles, headerStyles };
    }

    private _findDateXf(): number | null {
        const entry = this.zip.getEntry('xl/styles.bin');
        if (!entry) return null;
        // Reuse the styles parser (handles custom BrtFmt date formats, e.g.
        // xlspy's yyyy-mm-dd), instead of only built-in numFmt ranges.
        const reader = new BiffReaderWriter(entry.getData());
        while (reader.readStyles()) {
            // Consume; results accumulate on the reader.
        }
        const numFmtIds = reader._xfIndexToNumFmtId;
        const customDates = reader._customNumFmts;
        for (let i = 0; i < numFmtIds.length; i++) {
            const numFmtId = numFmtIds[i];
            if (
                (numFmtId >= 14 && numFmtId <= 22) ||
                (numFmtId >= 45 && numFmtId <= 47) ||
                customDates.has(numFmtId)
            ) {
                return i;
            }
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Row / cell serialization
    // ------------------------------------------------------------------

    private _buildRows(
        rows: CellValue[][],
        headers: string[] | undefined,
        dataStyles: Map<number, number>,
        headerStyles: Map<number, number>,
        dateXf: number | null
    ): Buffer {
        const chunks: Buffer[] = [];
        const maxCol = Math.max(rows.reduce((m, r) => Math.max(m, r.length), 0), headers?.length ?? 0, 0);
        const colLast = Math.max(0, maxCol - 1);

        const pushRow = (rowNum: number, cellBytes: Buffer[]): void => {
            // BrtRowHeader (0x00, len 25): rw(4) + zero(4) + flags(9) + colFirst(4) + colLast(4)
            const header = Buffer.alloc(27);
            header[0] = 0x00;
            header[1] = 25;
            header.writeInt32LE(rowNum, 2);
            header[10] = 0x2c;
            header[11] = 0x01;
            header[15] = 0x01;
            header.writeInt32LE(0, 19);
            header.writeInt32LE(colLast, 23);
            chunks.push(header);

            for (const cell of cellBytes) {
                chunks.push(cell);
            }
        };

        if (headers) {
            const cells: Buffer[] = [];
            for (let c = 0; c < headers.length; c++) {
                const style = headerStyles.get(c) ?? 0;
                cells.push(this._stringCellBytes(c, style, headers[c]));
            }
            pushRow(0, cells);
        }

        let rowNum = headers ? 1 : 0;
        for (const row of rows) {
            const cells: Buffer[] = [];
            for (let c = 0; c < row.length; c++) {
                const raw = row[c];
                if (raw === null || raw === undefined) continue;
                cells.push(this._valueCellBytes(raw, c, dataStyles, dateXf));
            }
            pushRow(rowNum, cells);
            rowNum++;
        }

        return Buffer.concat(chunks);
    }

    private _valueCellBytes(
        raw: CellValue,
        col: number,
        dataStyles: Map<number, number>,
        dateXf: number | null
    ): Buffer {
        const fmtString = getFormat(raw);
        const val = fmtString !== null ? unwrapCell(raw) : raw;
        const colStyle = dataStyles.get(col) ?? 0;

        if (typeof val === 'number') {
            if (Number.isFinite(val) && Number.isInteger(val) && val >= RK_INT_LOWER && val <= RK_INT_UPPER) {
                const payload = Buffer.alloc(12);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(colStyle, 4);
                payload.writeInt32LE((val << 2) | 2, 8);
                return buildRecord(0x0002, payload);
            }
            if (Number.isFinite(val)) {
                const payload = Buffer.alloc(16);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(colStyle, 4);
                payload.writeDoubleLE(val, 8);
                return buildRecord(0x0005, payload);
            }
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (typeof val === 'bigint') {
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (typeof val === 'boolean') {
            const payload = Buffer.alloc(9);
            payload.writeUInt32LE(col, 0);
            payload.writeUInt32LE(colStyle, 4);
            payload[8] = val ? 1 : 0;
            return buildRecord(0x0004, payload);
        }
        if (val instanceof Date) {
            const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
            if (Number.isFinite(oaDate)) {
                const style = colStyle > 0 ? colStyle : (dateXf !== null ? dateXf : 0);
                const payload = Buffer.alloc(16);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(style, 4);
                payload.writeDoubleLE(oaDate, 8);
                return buildRecord(0x0005, payload);
            }
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (val !== null && val !== undefined) {
            return this._stringCellBytes(col, colStyle, String(val));
        }
        return this._stringCellBytes(col, colStyle, '');
    }

    private _stringCellBytes(col: number, style: number, text: string): Buffer {
        if (this.sharedStringsBuf === null) {
            // Inline string (BrtCellSt): col(4) + xf(4) + cch(4) + utf16.
            const payload = Buffer.alloc(12 + text.length * 2);
            payload.writeUInt32LE(col, 0);
            payload.writeUInt32LE(style, 4);
            payload.writeUInt32LE(text.length, 8);
            payload.write(text, 12, text.length * 2, 'utf16le');
            return buildRecord(0x0006, payload);
        }

        let index = this.stringIndexMap.get(text);
        if (index === undefined) {
            index = this.sharedStringsValues.length;
            this.sharedStringsValues.push(text);
            this.stringIndexMap.set(text, index);
            this.sharedStringsUnique++;
            this._appendCount++;
        }
        this.sharedStringsTotal++;
        this.sharedStringsDirty = true;

        const payload = Buffer.alloc(12);
        payload.writeUInt32LE(col, 0);
        payload.writeUInt32LE(style, 4);
        payload.writeUInt32LE(index, 8);
        return buildRecord(0x0007, payload);
    }

    private _sharedStringRefCount(sheetBuf: Buffer): number {
        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        let count = 0;
        while (readRecord(sheetBuf, pos, rec)) {
            if (rec.id === 0x0007) count++;
            pos = rec.dataEnd;
        }
        return count;
    }

    private _patchSheetMetadataFile(
        filePath: string,
        originalSheet: Buffer,
        rowsEnd: number,
        delta: number,
        autoFilterRanges: number[],
        lastRowIdx: number,
        lastCol: number,
    ): void {
        const fd = fs.openSync(filePath, 'r+');
        try {
            const patchInt32 = (offset: number, value: number): void => {
                const bytes = Buffer.alloc(4);
                bytes.writeInt32LE(value, 0);
                fs.writeSync(fd, bytes, 0, bytes.length, offset);
            };

            for (const filterOffset of autoFilterRanges) {
                const newOffset = filterOffset >= rowsEnd ? filterOffset + delta : filterOffset;
                patchInt32(newOffset + 4, lastRowIdx);
                patchInt32(newOffset + 12, lastCol);
            }

            const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
            let pos = 0;
            while (readRecord(originalSheet, pos, rec)) {
                // BrtWsDim (0x94), not BrtSel (0x98): 16-byte payload
                // [rwFirst(4), rwLast(4), colFirst(4), colLast(4)].
                if (rec.id === 0x0094 && rec.len >= 16) {
                    const shift = rec.dataStart >= rowsEnd ? delta : 0;
                    patchInt32(rec.dataStart + shift + 4, lastRowIdx);
                    patchInt32(rec.dataStart + shift + 12, lastCol);
                    break;
                }
                pos = rec.dataEnd;
            }
        } finally {
            fs.closeSync(fd);
        }
    }

    // ------------------------------------------------------------------
    // Rows region detection
    // ------------------------------------------------------------------

    private _findRowsRegion(
        sheetBuf: Buffer
    ): { rowsStart: number; rowsEnd: number; autoFilterRanges: number[] } {
        const autoFilterRanges: number[] = [];
        let rowsStart = -1;
        let lastCellEnd = -1;
        let lastW25 = -1;

        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(sheetBuf, pos, rec)) {
            if (rec.id === 0x00) {
                if (rowsStart === -1) rowsStart = rec.headerStart;
            } else if (rec.id >= 0x01 && rec.id <= 0x0b) {
                lastCellEnd = rec.dataEnd;
            } else if (rec.id === 0x25 && rec.len === 6) {
                lastW25 = rec.headerStart;
            } else if (rec.id === 0x00a1 && rec.len >= 16) {
                autoFilterRanges.push(rec.dataStart);
            }
            pos = rec.dataEnd;
        }

        if (rowsStart === -1) {
            // Empty sheet: insert before the trailing wrapper (last 0x25) or at the end.
            const insertAt = lastW25 >= 0 ? lastW25 : sheetBuf.length;
            return { rowsStart: insertAt, rowsEnd: insertAt, autoFilterRanges };
        }

        // Replace only the row records. The row-block terminator records that Excel
        // writes after the last row (0x92/0x217/0x1dd/0x1dc) must be preserved,
        // otherwise Excel drops the last written row.
        const rowsEnd = lastCellEnd >= 0 ? lastCellEnd : (lastW25 >= 0 ? lastW25 : sheetBuf.length);

        return { rowsStart, rowsEnd, autoFilterRanges };
    }

    /**
     * Keep the sheet dimension record (0x94) in sync with the new data.
     * Excel uses its rwLast (payload offset 4) and colLast (payload offset 12)
     * fields to read the full row range; stale values make the last written row
     * invisible to Excel (and drop it from pivot refreshes).
     */
    private _patchDimension(sheetBuf: Buffer, lastRowIdx: number, lastCol: number): void {
        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(sheetBuf, pos, rec)) {
            if (rec.id === 0x0094 && rec.len >= 16) {
                sheetBuf.writeInt32LE(lastRowIdx, rec.dataStart + 4);
                sheetBuf.writeInt32LE(lastCol, rec.dataStart + 12);
                return;
            }
            pos = rec.dataEnd;
        }
    }

    // ------------------------------------------------------------------
    // Pivot cache integration
    // ------------------------------------------------------------------

    private _patchPivotCaches(sheetName: string, rwLast: number, colLast: number): void {
        for (const path of this.pivotCacheDefPaths) {
            const entry = this.zip.getEntry(path);
            if (!entry) continue;
            const buf = Buffer.from(entry.getData());

            const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
            let pos = 0;
            let patched = false;
            let refreshXfPatched = false;

            while (readRecord(buf, pos, rec)) {
                pos = rec.dataEnd;

                if (rec.id === 0x00b3 && rec.len >= 4 && !refreshXfPatched) {
                    // BrtBeginPivotCacheDefinition flags: bit 0x04 of payload offset 3 = refreshOnLoad.
                    if (rec.dataStart + 4 <= buf.length) {
                        const v = buf[rec.dataStart + 3];
                        if ((v & 0x04) === 0) {
                            buf[rec.dataStart + 3] = v | 0x04;
                            patched = true;
                        }
                    }
                    refreshXfPatched = true;
                }

                if (rec.id === 0x00bb && rec.len >= 23) {
                    // BrtPivotCacheSource: [3] + cch(4) + name(utf16) + rwFirst(4) rwLast(4) colFirst(4) colLast(4)
                    const cch = buf.readUInt32LE(rec.dataStart + 3);
                    const nameStart = rec.dataStart + 7;
                    const refStart = nameStart + cch * 2;
                    const name = readUtf16(buf, nameStart, cch);
                    if (name === sheetName && refStart + 16 <= buf.length) {
                        buf.writeInt32LE(rwLast, refStart + 4);
                        buf.writeInt32LE(colLast, refStart + 12);
                        patched = true;
                    }
                }
            }

            if (patched) {
                this.zip.updateFile(path, buf);
            }
        }
    }

    /**
     * Remove the stale calculation chain (xl/calcChain.bin). Excel rebuilds it
     * on open. Removes the part, its Content_Types override and its workbook
     * rels entry. XLSB tables are not patched (documented limitation).
     */
    private _removeCalcChain(): void {
        const calcPath = 'xl/calcChain.bin';
        if (!this.zip.getEntry(calcPath) && !this.stagedParts.has(calcPath)) return;
        try {
            this.zip.deleteFile(calcPath);
        } catch {
            // Best-effort.
        }
        const stagedCalc = this.stagedParts.get(calcPath);
        if (stagedCalc) {
            this.stagedParts.delete(calcPath);
            try {
                fs.rmSync(stagedCalc, { force: true });
            } catch {
                // Best-effort.
            }
        }

        // XLSB declares non-workbook .bin parts with an Override (the Default
        // "bin" content type is the workbook itself), so remove that too.
        const ctPath = '[Content_Types].xml';
        if (this.zip.getEntry(ctPath) || this.stagedParts.has(ctPath)) {
            try {
                const ct = this._entryData(ctPath).toString('utf8');
                const updated = ct.replace(
                    /<Override\b[^>]*PartName="\/xl\/calcChain\.bin"[^>]*\/?>/gi,
                    '',
                );
                if (updated !== ct) this._writePartText(ctPath, updated);
            } catch {
                // Best-effort.
            }
        }

        const relsPath = 'xl/_rels/workbook.bin.rels';
        if (this.zip.getEntry(relsPath) || this.stagedParts.has(relsPath)) {
            try {
                const rels = this._entryData(relsPath).toString('utf8');
                const relTagRe = /<Relationship\b[^>]*>/g;
                let m: RegExpExecArray | null;
                let out = '';
                let last = 0;
                let changed = false;
                while ((m = relTagRe.exec(rels)) !== null) {
                    const attrs = parseTagAttributes(m[0]);
                    const type = attrs['Type'] ?? attrs['type'] ?? '';
                    const target = attrs['Target'] ?? attrs['target'] ?? '';
                    if (type.endsWith('/calcChain') && resolveRelTarget(target) === 'xl/calcChain.bin') {
                        out += rels.substring(last, m.index);
                        last = m.index + m[0].length;
                        changed = true;
                    }
                }
                if (changed) {
                    out += rels.substring(last);
                    this._writePartText(relsPath, out);
                }
            } catch {
                // Best-effort.
            }
        }
    }

    /** Write an XML/text part to its staged file when present, else into the ZIP. */
    private _writePartText(name: string, text: string): void {
        const stagedPath = this.stagedParts.get(name);
        if (stagedPath) {
            fs.writeFileSync(stagedPath, text, 'utf8');
        } else {
            this.zip.updateFile(name, Buffer.from(text, 'utf8'));
        }
    }
}

export default XlsbUpdater;
