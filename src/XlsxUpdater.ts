import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { CellValue, getFormat, unwrapCell } from './Formats';
import { writeBufferAtomically } from './atomicFile';
import { writeAdmZipStreaming, StagedZipPart } from './streamingZip';
import { timeToOaDate } from './writerHelpers';
import {
    escapeXmlText,
    escapeXlsxText,
    columnIndexToLetter,
    columnLetterToIndex,
    parseSharedStringsXml,
    trimTrailingEmptyRows,
    parseTagAttributes,
    resolveRelTarget,
    unescapeXml,
} from './xmlUtils';

/**
 * Options for {@link XlsxUpdater.replaceSheetData}.
 */
export interface ReplaceSheetDataOptions {
    /**
     * Optional header row written into row 1 of the target sheet.
     * When omitted, row 1 is filled with the data rows.
     */
    headers?: string[];
    /**
     * Style strategy for the new cells:
     * - `'inherit'` (default): reuse the dominant cell style of each column from the
     *   existing data (and the original header style for row 1), so date/number
     *   formatting of the sheet is preserved. Falls back to General (0).
     * - `'general'`: all new cells use style 0 (General).
     */
    styleFallback?: 'inherit' | 'general';
}

/** A one-pass synchronous or asynchronous source of worksheet rows. */
export type RowSource =
    | Iterable<ReadonlyArray<CellValue>>
    | AsyncIterable<ReadonlyArray<CellValue>>;

interface PivotCacheDef {
    path: string;
    xml: string;
}

interface CellRange {
    c1: number;
    r1: number;
    c2: number;
    r2: number;
}

/** Parse an A1-style range ("A1", "A1:B10", "$A$1:$B$10"). */
function parseCellRange(ref: string): CellRange | null {
    const m = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(ref.trim());
    if (!m) return null;
    const c1 = columnLetterToIndex(m[1].toUpperCase());
    const r1 = parseInt(m[2], 10);
    const c2 = m[3] ? columnLetterToIndex(m[3].toUpperCase()) : c1;
    const r2 = m[4] ? parseInt(m[4], 10) : r1;
    if (!Number.isFinite(c1) || !Number.isFinite(r1) || !Number.isFinite(c2) || !Number.isFinite(r2)) {
        return null;
    }
    return {
        c1: Math.min(c1, c2),
        r1: Math.min(r1, r2),
        c2: Math.max(c1, c2),
        r2: Math.max(r1, r2),
    };
}

/**
 * Update the data of a worksheet inside an existing XLSX file without rebuilding
 * the workbook. Everything outside the target sheet's cell data — pivot tables
 * and their caches, other sheets, styles, themes, defined names — is preserved
 * byte-for-byte. This makes it possible to refresh a data sheet that other
 * Excel features (pivot tables, charts, formulas) are wired to.
 *
 * ```ts
 * const updater = new XlsxUpdater('report.xlsx');
 * updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
 * updater.save('report_new.xlsx'); // or updater.save() to overwrite in place
 * ```
 */
export class XlsxUpdater {
    private readonly zip: AdmZip;
    private readonly sourcePath: string;
    private temporaryDirectory: string | null = null;
    private readonly stagedParts: StagedZipPart = new Map();

    private sheetNameToPath: Map<string, string> = new Map();
    private pivotCacheDefs: PivotCacheDef[] = [];
    private pivotTableDefPaths: string[] = [];

    private sharedStringsXml: string | null = null;
    private sharedStringsValues: string[] = [];
    private sharedStringsOriginalLength: number = 0;
    private sharedStringsCount: number = 0;
    private sharedStringsPending: string[] = [];
    private sharedStringsDirty: boolean = false;
    private stringIndexMap: Map<string, number> = new Map();

    private _oaEpoch: number;

    /**
     * @param path Path to an existing `.xlsx` file.
     */
    constructor(filePath: string) {
        this.sourcePath = filePath;
        this._oaEpoch = Date.UTC(1899, 11, 30);

        if (!fs.existsSync(filePath)) {
            throw new Error(`XlsxUpdater: file not found: ${filePath}`);
        }

        this.zip = new AdmZip(filePath);

        if (this.zip.getEntry('xl/workbook.bin')) {
            throw new Error(
                'XlsxUpdater: XLSB files are not supported yet. Convert the workbook to XLSX first.'
            );
        }
        if (!this.zip.getEntry('xl/workbook.xml')) {
            throw new Error('XlsxUpdater: not a valid XLSX workbook (xl/workbook.xml missing).');
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
            throw new Error(`XlsxUpdater: sheet "${sheetName}" not found in the workbook.`);
        }

        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsxUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const sheetXml = this._entryData(sheetPath).toString('utf8');
        this._ensureSharedStringsLoaded();
        if (this.sharedStringsXml !== null) {
            this.sharedStringsCount = Math.max(
                0,
                this.sharedStringsCount - this._sharedStringRefCount(sheetXml),
            );
            this.sharedStringsDirty = true;
        }

        let colStyles: Map<number, number> = new Map();
        let headerStyles: Map<number, number> = new Map();
        let dateStyle: number | null = null;

        if (styleFallback !== 'general') {
            const collected = this._collectExistingStyles(sheetXml);
            colStyles = collected.dataStyles;
            headerStyles = collected.headerStyles;
            dateStyle = this._findDateStyleIndex();
        }

        const lastCol = Math.max(rows.reduce((m, r) => Math.max(m, r.length), 0), headers?.length ?? 0) - 1;
        const totalRows = rows.length + (headers ? 1 : 0);

        const newSheetData = this._buildSheetData(rows, headers, colStyles, headerStyles, dateStyle);
        const newXml = this._patchSheetXml(sheetXml, newSheetData, totalRows, lastCol);
        this.zip.updateFile(sheetPath, Buffer.from(newXml, 'utf8'));
        this._discardStagedPart(sheetPath);

        if (this.sharedStringsXml !== null) {
            this._commitSharedStrings();
        }

        const dimRef = totalRows > 0 && lastCol >= 0
            ? `A1:${columnIndexToLetter(lastCol)}${totalRows}`
            : 'A1';
        this._updatePivotCaches(sheetName, dimRef, rows.length);
        this._addRefreshOnLoadToPivotTables();
        this._removeCalcChain();
        this._updateTableRefs(sheetName, dimRef);
    }

    /**
     * Replace worksheet rows from a one-pass source without materialising the
     * complete result set. Use {@link saveStreaming} to keep the final ZIP
     * write streaming as well.
     */
    async replaceSheetDataStream(
        sheetName: string,
        rows: RowSource,
        options: ReplaceSheetDataOptions = {},
    ): Promise<void> {
        const { headers, styleFallback = 'inherit' } = options;
        if (styleFallback !== 'inherit' && styleFallback !== 'general') {
            throw new Error("XlsxUpdater: styleFallback must be 'inherit' or 'general'.");
        }

        const sheetPath = this.sheetNameToPath.get(sheetName);
        if (!sheetPath) {
            throw new Error(`XlsxUpdater: sheet "${sheetName}" not found in the workbook.`);
        }

        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsxUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const oldStage = this.stagedParts.get(sheetPath);
        const beforeValuesLength = this.sharedStringsValues.length;
        const beforePendingLength = this.sharedStringsPending.length;
        const beforeSharedStringsCount = this.sharedStringsCount;
        const beforeSharedStringsDirty = this.sharedStringsDirty;
        const beforeSharedStringsXml = this.sharedStringsXml;
        const beforeSharedStringsOriginalLength = this.sharedStringsOriginalLength;
        const beforeStringIndexMap = new Map(this.stringIndexMap);
        const sharedStringsEntry = this.zip.getEntry('xl/sharedStrings.xml');
        const beforeSharedStringsBytes = sharedStringsEntry || this.stagedParts.has('xl/sharedStrings.xml')
            ? Buffer.from(this._entryData('xl/sharedStrings.xml'))
            : null;
        const createdParts: string[] = [];
        const oldSheetXml = this._entryData(sheetPath).toString('utf8');
        const rowsPath = this._temporaryPath('.rows');
        const outputPath = this._temporaryPath('.xml');
        createdParts.push(rowsPath, outputPath);
        let output: fs.WriteStream | null = null;

        try {
            this._ensureSharedStringsLoaded();
            if (this.sharedStringsXml !== null) {
                this.sharedStringsCount = Math.max(
                    0,
                    this.sharedStringsCount - this._sharedStringRefCount(oldSheetXml),
                );
                this.sharedStringsDirty = true;
            }

            let colStyles = new Map<number, number>();
            let headerStyles = new Map<number, number>();
            let dateStyle: number | null = null;
            if (styleFallback !== 'general') {
                const collected = this._collectExistingStyles(oldSheetXml);
                colStyles = collected.dataStyles;
                headerStyles = collected.headerStyles;
                dateStyle = this._findDateStyleIndex();
            }

            const prefixMatch = /<([A-Za-z_][\w.-]*:)?sheetData\b/.exec(oldSheetXml);
            const prefix = prefixMatch?.[1] ?? '';
            const rowOutput = fs.createWriteStream(rowsPath);
            output = rowOutput;
            let outputError: Error | null = null;
            rowOutput.on('error', error => { outputError = error; });

            let nextRow = 1;
            let width = headers?.length ?? 0;
            let dataRowsSeen = 0;
            let keptDataRows = 0;
            let lastKeptEnd = 0;
            let pendingEmptyWidth = 0;
            let writtenBytes = 0;

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

            if (headers !== undefined) {
                const cells = headers.map((value, column) =>
                    this._valueCell(value, column, nextRow, headerStyles, dateStyle));
                const header = `<${prefix}row r="${nextRow}">${cells.join('')}</${prefix}row>`;
                await write(Buffer.from(header, 'utf8'));
                lastKeptEnd = writtenBytes;
                nextRow++;
            }

            for await (const sourceRow of rows) {
                if (typeof sourceRow === 'string' || sourceRow instanceof Uint8Array) {
                    throw new TypeError('Each row must be an iterable of cells, not text');
                }
                const row = Array.from(sourceRow) as CellValue[];
                const cells = row.map((value, column) =>
                    value === null || value === undefined
                        ? ''
                        : this._valueCell(value, column, nextRow, colStyles, dateStyle));
                const rowXml = `<${prefix}row r="${nextRow}">${cells.join('')}</${prefix}row>`;
                await write(Buffer.from(rowXml, 'utf8'));
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

            const totalRows = keptDataRows + (headers !== undefined ? 1 : 0);
            const lastCol = width - 1;
            const dimRef = totalRows > 0 && lastCol >= 0
                ? `A1:${columnIndexToLetter(lastCol)}${totalRows}`
                : 'A1';
            this._patchSheetFile(oldSheetXml, rowsPath, outputPath, dimRef);
            this._stagePart(sheetPath, outputPath);
            this._commitSharedStrings();
            this._updatePivotCaches(sheetName, dimRef, keptDataRows);
            this._addRefreshOnLoadToPivotTables();
            this._removeCalcChain();
            this._updateTableRefs(sheetName, dimRef);

            if (oldStage && oldStage !== outputPath) {
                this._removeTemporaryFile(oldStage);
            }
            createdParts.splice(createdParts.indexOf(outputPath), 1);
        } catch (error) {
            this.sharedStringsValues.length = beforeValuesLength;
            this.sharedStringsPending.length = beforePendingLength;
            this.sharedStringsCount = beforeSharedStringsCount;
            this.sharedStringsDirty = beforeSharedStringsDirty;
            this.sharedStringsXml = beforeSharedStringsXml;
            this.sharedStringsOriginalLength = beforeSharedStringsOriginalLength;
            this.stringIndexMap.clear();
            for (const [value, index] of beforeStringIndexMap) this.stringIndexMap.set(value, index);
            if (beforeSharedStringsBytes) {
                this.zip.updateFile('xl/sharedStrings.xml', beforeSharedStringsBytes);
            }
            if (oldStage) this.stagedParts.set(sheetPath, oldStage);
            else this.stagedParts.delete(sheetPath);
            throw error;
        } finally {
            if (output && !output.closed) output.destroy();
            for (const temporaryPath of createdParts) {
                this._removeTemporaryFile(temporaryPath);
            }
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

        // Keep the updater reusable after an in-place save. The synchronous
        // AdmZip path is refreshed with the staged parts only after the output
        // archive has been safely installed.
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
        if (!entry) throw new Error(`XlsxUpdater: ZIP member missing: ${name}`);
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
        const relsEntry = this.zip.getEntry('xl/_rels/workbook.xml.rels');
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

        const wbEntry = this.zip.getEntry('xl/workbook.xml');
        const wbXml = wbEntry ? wbEntry.getData().toString('utf8') : '';
        const sheetRe = /<sheet\b[^>]*>/g;
        let match;
        while ((match = sheetRe.exec(wbXml)) !== null) {
            const tag = match[0];
            const attrs = parseTagAttributes(tag);
            const rawName = attrs['name'];
            const rId = attrs['r:id'] ?? attrs['id'] ?? attrs['rId'];
            if (rawName === undefined || rId === undefined) continue;
            const name = unescapeXml(rawName);
            const target = rIdToTarget[rId] ?? '';
            const fullPath = resolveRelTarget(target);
            if (fullPath) {
                this.sheetNameToPath.set(name, fullPath);
            }
        }
        // 1904 date system for Date→OA conversion.
        const wbPrMatch = /<workbookPr\b[^>]*>/i.exec(wbXml);
        if (wbPrMatch) {
            const wbPrAttrs = parseTagAttributes(wbPrMatch[0]);
            const v = (wbPrAttrs['date1904'] ?? '').trim().toLowerCase();
            this._oaEpoch =
                (v === '1' || v === 'true') ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
        }

        for (const entry of this.zip.getEntries()) {
            const name = entry.entryName;
            if (/^xl\/pivotCache\/pivotCacheDefinition\d*\.xml$/.test(name)) {
                this.pivotCacheDefs.push({ path: name, xml: entry.getData().toString('utf8') });
            } else if (/^xl\/pivotTables\/pivotTable\d*\.xml$/.test(name)) {
                this.pivotTableDefPaths.push(name);
            }
        }
    }

    // ------------------------------------------------------------------
    // Shared strings
    // ------------------------------------------------------------------

    private _ensureSharedStringsLoaded(): void {
        if (this.stringIndexMap.size > 0 || this.sharedStringsXml !== null) {
            return;
        }
        const entry = this.zip.getEntry('xl/sharedStrings.xml');
        if (!entry && !this.stagedParts.has('xl/sharedStrings.xml')) {
            this.sharedStringsXml = null;
            return;
        }
        this.sharedStringsXml = this._entryData('xl/sharedStrings.xml').toString('utf8');
        this.sharedStringsValues = parseSharedStringsXml(this.sharedStringsXml);

        const countMatch = /<(?:[A-Za-z_][\w.-]*:)?sst\b[^>]*>/i.exec(this.sharedStringsXml);
        const countAttr = countMatch ? /(?:^|\s)count\s*=\s*["'](\d+)["']/.exec(countMatch[0]) : null;
        this.sharedStringsCount = countAttr ? parseInt(countAttr[1], 10) : this.sharedStringsValues.length;

        this.sharedStringsValues.forEach((value, index) => {
            if (!this.stringIndexMap.has(value)) {
                this.stringIndexMap.set(value, index);
            }
        });
        this.sharedStringsOriginalLength = this.sharedStringsValues.length;
        this.sharedStringsPending = [];
        this.sharedStringsDirty = false;
    }

    private _commitSharedStrings(): void {
        if (this.sharedStringsXml === null || !this.sharedStringsDirty) {
            return;
        }

        let xml = this.sharedStringsXml;

        const sstMatch = /<(?<prefix>[A-Za-z_][\w.-]*:)?sst\b[^>]*>/i.exec(xml);
        if (!sstMatch || sstMatch.index === undefined) {
            throw new Error('XlsxUpdater: sharedStrings.xml has no sst element.');
        }
        const tag = sstMatch[0];
        const prefix = sstMatch.groups?.prefix ?? '';
        const newTag = XlsxUpdater._replaceOrAddXmlAttribute(
            XlsxUpdater._replaceOrAddXmlAttribute(tag, 'count', String(this.sharedStringsCount)),
            'uniqueCount',
            String(this.sharedStringsValues.length),
        );
        xml = xml.substring(0, sstMatch.index) + newTag + xml.substring(sstMatch.index + tag.length);

        const closingRe = new RegExp(`</${prefix}sst\\s*>\\s*$`, 'i');
        const closingMatch = closingRe.exec(xml);
        if (!closingMatch || closingMatch.index === undefined) {
            throw new Error('XlsxUpdater: sharedStrings.xml has no closing sst element.');
        }

        const appended = this.sharedStringsPending.map(txt => {
            const clean = escapeXlsxText(txt);
            const preserve = clean.length > 0 && (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean));
            const tAttr = preserve ? ' xml:space="preserve"' : '';
            return `<${prefix}si><${prefix}t${tAttr}>${clean}</${prefix}t></${prefix}si>`;
        }).join('');

        xml = xml.substring(0, closingMatch.index) + appended + xml.substring(closingMatch.index);
        this.zip.updateFile('xl/sharedStrings.xml', Buffer.from(xml, 'utf8'));
        this.sharedStringsXml = xml;
        this.sharedStringsOriginalLength = this.sharedStringsValues.length;
        this.sharedStringsPending = [];
        this.sharedStringsDirty = false;
    }

    // ------------------------------------------------------------------
    // Styles
    // ------------------------------------------------------------------

    private _collectExistingStyles(sheetXml: string): { dataStyles: Map<number, number>; headerStyles: Map<number, number> } {
        const dataCounts = new Map<number, Map<number, number>>();
        const headerStyles = new Map<number, number>();

        const sdStart = sheetXml.indexOf('<sheetData>');
        if (sdStart === -1) return { dataStyles: new Map(), headerStyles };
        const sdEnd = sheetXml.indexOf('</sheetData>', sdStart);
        if (sdEnd === -1) return { dataStyles: new Map(), headerStyles };

        const region = sheetXml.substring(sdStart, sdEnd);
        // `r` on a cell is optional: a cell without it sits right after the previous
        // cell of its row (column A when first), so track the position per row.
        const cellRe = /<row\b[^>]*>|<c\b[^>]*>/g;
        let m;
        let row = 0;
        let nextCol = 0;
        while ((m = cellRe.exec(region)) !== null) {
            const tag = m[0];
            if (tag.startsWith('<row')) {
                const rowAttr = /\br="(\d+)"/.exec(tag);
                row = rowAttr ? parseInt(rowAttr[1], 10) : row + 1;
                nextCol = 0;
                continue;
            }
            const rMatch = /r="([A-Z]+)(\d+)"/.exec(tag);
            let col = nextCol;
            if (rMatch) {
                col = columnLetterToIndex(rMatch[1]);
                row = parseInt(rMatch[2], 10);
            }
            nextCol = col + 1;
            const sMatch = /s="(\d+)"/.exec(tag);
            if (!sMatch) continue;
            const s = parseInt(sMatch[1], 10);
            if (s <= 0) continue;

            if (row === 1) {
                if (!headerStyles.has(col)) headerStyles.set(col, s);
                continue;
            }
            let colCounts = dataCounts.get(col);
            if (!colCounts) {
                colCounts = new Map();
                dataCounts.set(col, colCounts);
            }
            colCounts.set(s, (colCounts.get(s) ?? 0) + 1);
        }

        const dataStyles = new Map<number, number>();
        for (const [col, colCounts] of dataCounts) {
            let best = 0;
            let bestCount = -1;
            for (const [s, count] of colCounts) {
                if (count > bestCount) {
                    best = s;
                    bestCount = count;
                }
            }
            if (bestCount > 0) dataStyles.set(col, best);
        }

        return { dataStyles, headerStyles };
    }

    private _findDateStyleIndex(): number | null {
        const entry = this.zip.getEntry('xl/styles.xml');
        if (!entry) return null;
        const xml = entry.getData().toString('utf8');

        const customDateFormats = new Set<number>();
        const numFmtRe = /<numFmt\s+[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
        let m;
        while ((m = numFmtRe.exec(xml)) !== null) {
            const id = parseInt(m[1], 10);
            const code = m[2].toLowerCase();
            if (code.includes('yy') || code.includes('mm') || code.includes('dd') || code.includes('h:mm')) {
                customDateFormats.add(id);
            }
        }

        const cellXfsStart = xml.indexOf('<cellXfs');
        if (cellXfsStart === -1) return null;
        const cellXfsEnd = xml.indexOf('</cellXfs>', cellXfsStart);
        if (cellXfsEnd === -1) return null;
        const cellXfsContent = xml.substring(cellXfsStart, cellXfsEnd);

        const xfRe = /<xf\b[^>]*>/g;
        let index = 0;
        while ((m = xfRe.exec(cellXfsContent)) !== null) {
            const nfMatch = /numFmtId="(\d+)"/.exec(m[0]);
            const nfId = nfMatch ? parseInt(nfMatch[1], 10) : 0;
            const isDate = (nfId >= 14 && nfId <= 22) ||
                (nfId >= 45 && nfId <= 47) ||
                customDateFormats.has(nfId);
            if (isDate) return index;
            index++;
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Cell / row serialization
    // ------------------------------------------------------------------

    private _buildSheetData(
        rows: CellValue[][],
        headers: string[] | undefined,
        colStyles: Map<number, number>,
        headerStyles: Map<number, number>,
        dateStyle: number | null
    ): string {
        let out = '';
        let rowNum = 0;

        if (headers) {
            rowNum = 1;
            out += '<row r="1">';
            for (let c = 0; c < headers.length; c++) {
                const style = headerStyles.get(c) ?? 0;
                out += this._stringCell(headers[c], c, rowNum, style);
            }
            out += '</row>';
        }

        for (const row of rows) {
            rowNum++;
            out += `<row r="${rowNum}">`;
            for (let c = 0; c < row.length; c++) {
                const raw = row[c];
                if (raw === null || raw === undefined) continue;
                out += this._valueCell(raw, c, rowNum, colStyles, dateStyle);
            }
            out += '</row>';
        }
        return out;
    }

    private _valueCell(
        raw: CellValue,
        col: number,
        rowNum: number,
        colStyles: Map<number, number>,
        dateStyle: number | null
    ): string {
        const colRef = columnIndexToLetter(col);
        const fmtString = getFormat(raw);
        const val = fmtString !== null ? unwrapCell(raw) : raw;
        const colStyle = colStyles.get(col) ?? 0;

        if (typeof val === 'number') {
            if (Number.isFinite(val)) {
                return `<c r="${colRef}${rowNum}"${colStyle > 0 ? ` s="${colStyle}"` : ''}><v>${val}</v></c>`;
            }
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (typeof val === 'bigint') {
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (typeof val === 'boolean') {
            return `<c r="${colRef}${rowNum}" t="b"${colStyle > 0 ? ` s="${colStyle}"` : ''}><v>${val ? 1 : 0}</v></c>`;
        }
        if (val instanceof Date) {
            const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
            if (Number.isFinite(oaDate)) {
                const style = colStyle > 0 ? colStyle : (dateStyle !== null ? dateStyle : 0);
                return `<c r="${colRef}${rowNum}"${style > 0 ? ` s="${style}"` : ''}><v>${oaDate}</v></c>`;
            }
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (val !== null && val !== undefined) {
            return this._stringCell(String(val), col, rowNum, colStyle);
        }
        return '';
    }

    private _stringCell(text: string, col: number, rowNum: number, style: number): string {
        const colRef = columnIndexToLetter(col);
        const styleAttr = style > 0 ? ` s="${style}"` : '';

        if (this.sharedStringsXml === null) {
            const clean = escapeXlsxText(text);
            const preserve = clean.length > 0 && (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean));
            const tAttr = preserve ? ' xml:space="preserve"' : '';
            return `<c r="${colRef}${rowNum}" t="inlineStr"${styleAttr}><is><t${tAttr}>${clean}</t></is></c>`;
        }

        let index = this.stringIndexMap.get(text);
        if (index === undefined) {
            index = this.sharedStringsValues.length;
            this.sharedStringsValues.push(text);
            this.stringIndexMap.set(text, index);
            this.sharedStringsPending.push(text);
        }
        this.sharedStringsCount++;
        this.sharedStringsDirty = true;
        return `<c r="${colRef}${rowNum}" t="s"${styleAttr}><v>${index}</v></c>`;
    }

    private static _replaceOrAddXmlAttribute(tag: string, name: string, value: string): string {
        const escaped = escapeXmlText(value);
        const attribute = new RegExp(`(\\s${name}\\s*=\\s*["'])[^"']*(["'])`, 'i');
        if (attribute.test(tag)) {
            return tag.replace(attribute, `$1${escaped}$2`);
        }
        const insertion = ` ${name}="${escaped}"`;
        return tag.endsWith('/>')
            ? `${tag.slice(0, -2)}${insertion}/>`
            : `${tag.slice(0, -1)}${insertion}>`;
    }

    private _sharedStringRefCount(sheetXml: string): number {
        const cellTags = sheetXml.match(/<(?:(?:[A-Za-z_][\w.-]*):)?c\b[^>]*>/gi) ?? [];
        return cellTags.filter(tag => /(?:^|\s)t\s*=\s*["']s["']/i.test(tag)).length;
    }

    // ------------------------------------------------------------------
    // Worksheet XML patching
    // ------------------------------------------------------------------

    private _patchSheetXml(sheetXml: string, newSheetData: string, totalRows: number, lastCol: number): string {
        let xml = sheetXml;

        const dimRef = totalRows > 0 && lastCol >= 0
            ? `A1:${columnIndexToLetter(lastCol)}${totalRows}`
            : 'A1';

        // Update the dimension element, or insert one if missing.
        const dimRe = /<dimension\b[^>]*>/;
        const dimMatch = dimRe.exec(xml);
        if (dimMatch) {
            xml = xml.substring(0, dimMatch.index) +
                dimMatch[0].replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(dimMatch.index + dimMatch[0].length);
        } else {
            const rootRe = /<worksheet\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (rootMatch) {
                const insertAt = rootMatch.index + rootMatch[0].length;
                xml = xml.substring(0, insertAt) + `<dimension ref="${dimRef}"/>` + xml.substring(insertAt);
            }
        }

        // Replace the entire sheetData element content.
        const sdOpen = '<sheetData>';
        const sdStart = xml.indexOf(sdOpen);
        if (sdStart === -1) {
            const selfClosing = /<sheetData\s*\/\s*>/;
            const scMatch = selfClosing.exec(xml);
            if (scMatch) {
                xml = xml.substring(0, scMatch.index) +
                    '<sheetData>' + newSheetData + '</sheetData>' +
                    xml.substring(scMatch.index + scMatch[0].length);
                return xml;
            }
            throw new Error('XlsxUpdater: <sheetData> element not found in worksheet XML.');
        }
        const sdEnd = xml.indexOf('</sheetData>', sdStart);
        if (sdEnd === -1) {
            throw new Error('XlsxUpdater: <sheetData> element not closed in worksheet XML.');
        }
        xml = xml.substring(0, sdStart + sdOpen.length) + newSheetData + xml.substring(sdEnd);

        // Keep the autoFilter range in sync with the new data bounds.
        const afRe = /<autoFilter\b[^>]*>/;
        const afMatch = afRe.exec(xml);
        if (afMatch && totalRows > 0 && lastCol >= 0 && afMatch[0].includes('ref=')) {
            xml = xml.substring(0, afMatch.index) +
                afMatch[0].replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(afMatch.index + afMatch[0].length);
        }

        return xml;
    }

    private _patchSheetFile(
        sheetXml: string,
        rowsPath: string,
        outputPath: string,
        dimensionRef: string,
    ): void {
        const source = Buffer.from(sheetXml, 'utf8');
        const text = sheetXml;
        const replacements: Array<{
            start: number;
            end: number;
            data?: Buffer;
            rows?: boolean;
            suffix?: Buffer;
        }> = [];

        const emptySheetData = /<(?<prefix>[A-Za-z_][\w.-]*:)?sheetData\b(?<attrs>[^>]*)\/>/i.exec(text);
        if (emptySheetData && emptySheetData.index !== undefined) {
            const prefix = emptySheetData.groups?.prefix ?? '';
            replacements.push({
                start: emptySheetData.index,
                end: emptySheetData.index + emptySheetData[0].length,
                data: Buffer.from(`<${prefix}sheetData${emptySheetData.groups?.attrs ?? ''}>`, 'utf8'),
                rows: true,
                suffix: Buffer.from(`</${prefix}sheetData>`, 'utf8'),
            });
        } else {
            const opening = /<(?<prefix>[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/i.exec(text);
            if (!opening || opening.index === undefined) {
                throw new Error('XlsxUpdater: <sheetData> element not found in worksheet XML.');
            }
            const prefix = opening.groups?.prefix ?? '';
            const closeRe = new RegExp(`</${prefix}sheetData\\s*>`, 'i');
            const closing = closeRe.exec(text.slice(opening.index + opening[0].length));
            if (!closing || closing.index === undefined) {
                throw new Error('XlsxUpdater: <sheetData> element not closed in worksheet XML.');
            }
            const closingStart = opening.index + opening[0].length + closing.index;
            replacements.push({
                start: opening.index + opening[0].length,
                end: closingStart,
                rows: true,
            });
        }

        const dimension = /<(?<prefix>[A-Za-z_][\w.-]*:)?dimension\b[^>]*>/i.exec(text);
        if (dimension && dimension.index !== undefined) {
            replacements.push({
                start: dimension.index,
                end: dimension.index + dimension[0].length,
                data: Buffer.from(
                    XlsxUpdater._replaceOrAddXmlAttribute(dimension[0], 'ref', dimensionRef),
                ),
            });
        } else {
            const root = /<(?<prefix>[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>/i.exec(text);
            if (root && root.index !== undefined) {
                const prefix = root.groups?.prefix ?? '';
                const insertAt = root.index + root[0].length;
                replacements.push({
                    start: insertAt,
                    end: insertAt,
                    data: Buffer.from(`<${prefix}dimension ref="${dimensionRef}"/>`, 'utf8'),
                });
            }
        }

        const autoFilter = /<(?<prefix>[A-Za-z_][\w.-]*:)?autoFilter\b[^>]*>/i.exec(text);
        if (autoFilter && autoFilter.index !== undefined) {
            replacements.push({
                start: autoFilter.index,
                end: autoFilter.index + autoFilter[0].length,
                data: Buffer.from(
                    XlsxUpdater._replaceOrAddXmlAttribute(autoFilter[0], 'ref', dimensionRef),
                    'utf8',
                ),
            });
        }

        replacements.sort((a, b) => a.start - b.start || a.end - b.end);
        const fd = fs.openSync(outputPath, 'w');
        let cursor = 0;
        const writeBuffer = (data: Buffer) => {
            fs.writeSync(fd, data);
        };
        const copyRows = () => {
            const rowsFd = fs.openSync(rowsPath, 'r');
            try {
                const chunk = Buffer.alloc(1024 * 1024);
                const length = fs.statSync(rowsPath).size;
                let position = 0;
                while (position < length) {
                    const count = fs.readSync(rowsFd, chunk, 0, Math.min(chunk.length, length - position), position);
                    if (count <= 0) throw new Error('XlsxUpdater: failed to read staged rows.');
                    fs.writeSync(fd, chunk, 0, count);
                    position += count;
                }
            } finally {
                fs.closeSync(rowsFd);
            }
        };
        try {
            for (const replacement of replacements) {
                const startByte = Buffer.byteLength(text.slice(0, replacement.start), 'utf8');
                const endByte = Buffer.byteLength(text.slice(0, replacement.end), 'utf8');
                if (startByte < cursor) throw new Error('XlsxUpdater: overlapping worksheet replacements.');
                writeBuffer(source.subarray(cursor, startByte));
                if (replacement.data) writeBuffer(replacement.data);
                if (replacement.rows) copyRows();
                if (replacement.suffix) writeBuffer(replacement.suffix);
                cursor = endByte;
            }
            writeBuffer(source.subarray(cursor));
        } finally {
            fs.closeSync(fd);
        }
    }

    // ------------------------------------------------------------------
    // Pivot table integration
    // ------------------------------------------------------------------

    private _updatePivotCaches(sheetName: string, dimRef: string, recordCount: number): void {
        for (const cache of this.pivotCacheDefs) {
            const wsRe = /<worksheetSource\b[^>]*>/;
            const wsMatch = wsRe.exec(cache.xml);
            if (!wsMatch) continue;
            const wsTag = wsMatch[0];
            const sheetAttr = /sheet="([^"]*)"/.exec(wsTag);
            if (!sheetAttr || sheetAttr[1] !== sheetName) continue;

            let xml = cache.xml;
            xml = xml.substring(0, wsMatch.index) +
                wsTag.replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(wsMatch.index + wsTag.length);

            const rootRe = /<pivotCacheDefinition\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (rootMatch) {
                let rootTag = rootMatch[0];
                if (/recordCount="\d+"/.test(rootTag)) {
                    rootTag = rootTag.replace(/recordCount="\d+"/, `recordCount="${recordCount}"`);
                } else if (!/\/>$/.test(rootTag)) {
                    rootTag = rootTag.replace(/>$/, ` recordCount="${recordCount}">`);
                }
                xml = xml.substring(0, rootMatch.index) + rootTag + xml.substring(rootMatch.index + rootMatch[0].length);
            }

            this.zip.updateFile(cache.path, Buffer.from(xml, 'utf8'));
        }
    }

    private _addRefreshOnLoadToPivotTables(): void {
        for (const path of this.pivotTableDefPaths) {
            const entry = this.zip.getEntry(path);
            if (!entry) continue;
            let xml = entry.getData().toString('utf8');

            const rootRe = /<pivotTableDefinition\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (!rootMatch || rootMatch[0].includes('refreshOnLoad')) continue;

            const newTag = rootMatch[0].replace(/(\/?>)$/, ' refreshOnLoad="1"$1');
            xml = xml.substring(0, rootMatch.index) + newTag + xml.substring(rootMatch.index + rootMatch[0].length);
            this.zip.updateFile(path, Buffer.from(xml, 'utf8'));
        }
    }

    /**
     * Remove the stale calculation chain (xl/calcChain.xml). Excel rebuilds it
     * on open; keeping stale <c> entries for the replaced sheet would force a
     * repair prompt. Removes the part, its Content_Types override and its
     * workbook rels entry (xlspy parity).
     */
    private _removeCalcChain(): void {
        const calcPath = 'xl/calcChain.xml';
        if (!this.zip.getEntry(calcPath) && !this.stagedParts.has(calcPath)) {
            this._ensureCalcPrFullCalc();
            return;
        }
        try {
            this.zip.deleteFile(calcPath);
        } catch {
            // Best-effort; continue with rels/content-type cleanup.
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

        const ctEntry = this.zip.getEntry('[Content_Types].xml');
        if (ctEntry || this.stagedParts.has('[Content_Types].xml')) {
            try {
                let ct = this._entryData('[Content_Types].xml').toString('utf8');
                const before = ct;
                ct = ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/?>/gi, '');
                if (ct !== before) {
                    this._writePartText('[Content_Types].xml', ct);
                }
            } catch {
                // Best-effort.
            }
        }

        const relsPath = 'xl/_rels/workbook.xml.rels';
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
                    const isCalcChain = type.endsWith('/calcChain') &&
                        resolveRelTarget(target) === 'xl/calcChain.xml';
                    if (isCalcChain) {
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
        this._ensureCalcPrFullCalc();
    }

    private _writePartText(name: string, text: string): void {
        if (this.stagedParts.has(name)) {
            fs.writeFileSync(this.stagedParts.get(name)!, text, 'utf8');
        } else {
            this.zip.updateFile(name, Buffer.from(text, 'utf8'));
        }
    }

    private _ensureCalcPrFullCalc(): void {
        try {
            const wbPath = 'xl/workbook.xml';
            if (!this.zip.getEntry(wbPath) && !this.stagedParts.has(wbPath)) return;
            let xml = this._entryData(wbPath).toString('utf8');
            const calcPrRe = /<(?:[A-Za-z_][\w.-]*:)?calcPr\b[^>]*>/i;
            const match = calcPrRe.exec(xml);
            if (match && match.index !== undefined) {
                const newTag = XlsxUpdater._replaceOrAddXmlAttribute(match[0], 'fullCalcOnLoad', '1');
                // The replacement can change the tag length, so consume the
                // *original* tag (`match[0]`), not `newTag`.
                xml = xml.substring(0, match.index) + newTag + xml.substring(match.index + match[0].length);
            } else {
                const insertAt = XlsxUpdater._calcPrInsertPosition(xml);
                if (insertAt < 0) return;
                xml = xml.substring(0, insertAt) + '<calcPr fullCalcOnLoad="1"/>' + xml.substring(insertAt);
            }
            this._writePartText(wbPath, xml);
        } catch {
            // Best-effort.
        }
    }

    /**
     * Insertion offset for a missing `<calcPr>`. The element follows
     * sheets/externalReferences/definedNames in the workbook schema, so insert
     * after the last of those that exists instead of right after `<workbook>`.
     */
    private static _calcPrInsertPosition(xml: string): number {
        const prefixed = '(?:[A-Za-z_][\\w.-]*:)?';
        const candidates = [
            new RegExp(`</${prefixed}sheets\\s*>`, 'i'),
            new RegExp(`<${prefixed}externalReferences\\b[^>]*/>`, 'i'),
            new RegExp(`<${prefixed}externalReferences\\b[^>]*>[\\s\\S]*?</${prefixed}externalReferences\\s*>`, 'i'),
            new RegExp(`<${prefixed}definedNames\\b[^>]*/>`, 'i'),
            new RegExp(`<${prefixed}definedNames\\b[^>]*>[\\s\\S]*?</${prefixed}definedNames\\s*>`, 'i'),
        ];
        let insertAt = -1;
        for (const re of candidates) {
            const m = re.exec(xml);
            if (m && m.index !== undefined) {
                insertAt = Math.max(insertAt, m.index + m[0].length);
            }
        }
        if (insertAt >= 0) return insertAt;
        // Fallbacks for unusual workbooks: before the close tag, else after the root.
        const close = new RegExp(`</${prefixed}workbook\\s*>`, 'i').exec(xml);
        if (close && close.index !== undefined) return close.index;
        const root = new RegExp(`<${prefixed}workbook\\b[^>]*>`, 'i').exec(xml);
        if (!root || root.index === undefined) return -1;
        return root.index + root[0].length;
    }

    /**
     * Update ListObject table ranges (xl/tables/*.xml) for the replaced sheet.
     * Only tables anchored at A1 are updated to the new dimension; tables
     * elsewhere are left untouched (moving them would corrupt the sheet).
     * Empty replacements (dimRef "A1") do not shrink multi-column tables.
     */
    private _updateTableRefs(sheetName: string, dimRef: string): void {
        const sheetPath = this.sheetNameToPath.get(sheetName);
        if (!sheetPath) return;
        const slash = sheetPath.lastIndexOf('/');
        const dir = slash >= 0 ? sheetPath.substring(0, slash + 1) : '';
        const base = slash >= 0 ? sheetPath.substring(slash + 1) : sheetPath;
        const sheetRelsPath = `${dir}_rels/${base}.rels`;
        const sheetRelsEntry = this.zip.getEntry(sheetRelsPath);
        const hasStagedRels = this.stagedParts.has(sheetRelsPath);
        if (!sheetRelsEntry && !hasStagedRels) return;

        const tableTargets: string[] = [];
        try {
            const relsXml = this._entryData(sheetRelsPath).toString('utf8');
            const relTagRe = /<Relationship\b[^>]*>/g;
            let m: RegExpExecArray | null;
            while ((m = relTagRe.exec(relsXml)) !== null) {
                const attrs = parseTagAttributes(m[0]);
                const type = attrs['Type'] ?? attrs['type'] ?? '';
                const target = attrs['Target'] ?? attrs['target'] ?? '';
                if (!type.includes('/table') || !target) continue;
                let tablePath: string;
                if (target.startsWith('/')) {
                    tablePath = resolveRelTarget(target);
                } else {
                    // Resolve relative to the worksheet directory (e.g. ../tables/table1.xml).
                    const combined = dir + target;
                    const parts: string[] = [];
                    for (const seg of combined.split('/')) {
                        if (seg === '..') parts.pop();
                        else if (seg !== '.' && seg !== '') parts.push(seg);
                    }
                    tablePath = parts.join('/');
                }
                tableTargets.push(tablePath);
            }
        } catch {
            return;
        }
        if (tableTargets.length === 0) return;

        // Read every table first: expanding an A1 table must not overlap a table
        // that sits below it in the same columns.
        const tables: Array<{
            path: string;
            xml: string;
            tag: string;
            tagIndex: number;
            ref: string;
            range: CellRange | null;
            staged: boolean;
        }> = [];
        for (const tablePath of tableTargets) {
            try {
                const hasEntry = !!this.zip.getEntry(tablePath);
                const hasStaged = this.stagedParts.has(tablePath);
                if (!hasEntry && !hasStaged) continue;
                const xml = this._entryData(tablePath).toString('utf8');
                const tableTagRe = /<(?:[A-Za-z_][\w.-]*:)?table\b[^>]*>/i;
                const tableMatch = tableTagRe.exec(xml);
                if (!tableMatch || tableMatch.index === undefined) continue;
                const refMatch = /\bref\s*=\s*("([^"]*)"|'([^']*)')/i.exec(tableMatch[0]);
                const ref = refMatch ? (refMatch[2] ?? refMatch[3] ?? '') : '';
                tables.push({
                    path: tablePath,
                    xml,
                    tag: tableMatch[0],
                    tagIndex: tableMatch.index,
                    ref,
                    range: parseCellRange(ref),
                    staged: hasStaged,
                });
            } catch {
                // Best-effort per table.
            }
        }
        if (tables.length === 0) return;

        const dimRange = parseCellRange(dimRef);
        if (!dimRange) return;

        for (const table of tables) {
            try {
                const origRange = table.range;
                if (!origRange) continue;
                // Only A1-anchored tables track the sheet data range.
                if (origRange.r1 !== 1 || origRange.c1 !== 0) continue;
                // Don't collapse a multi-column table to a single cell.
                if (dimRef === 'A1' && !/^A1$/i.test(table.ref)) continue;
                const origCols = origRange.c2 - origRange.c1 + 1;
                const newCols = dimRange.c2 - dimRange.c1 + 1;
                // Changing the width would also require rebuilding <tableColumns>.
                if (newCols !== origCols) continue;

                // Cap at the first other table below this one so tables never overlap.
                let lastRow = dimRange.r2;
                for (const other of tables) {
                    if (other === table || !other.range) continue;
                    if (other.range.r1 <= 1) continue;
                    if (other.range.c1 > origRange.c2 || other.range.c2 < origRange.c1) continue;
                    lastRow = Math.min(lastRow, other.range.r1 - 1);
                }
                if (lastRow < 1) continue;

                const newRef = origRange.c2 === 0 && lastRow === 1
                    ? 'A1'
                    : `A1:${columnIndexToLetter(origRange.c2)}${lastRow}`;
                let xml = table.xml;
                const newTag = XlsxUpdater._replaceOrAddXmlAttribute(table.tag, 'ref', newRef);
                xml = xml.substring(0, table.tagIndex) + newTag +
                    xml.substring(table.tagIndex + table.tag.length);
                const afRe = /<(?:[A-Za-z_][\w.-]*:)?autoFilter\b[^>]*>/i;
                const afMatch = afRe.exec(xml);
                if (afMatch && afMatch.index !== undefined) {
                    const newAf = XlsxUpdater._replaceOrAddXmlAttribute(afMatch[0], 'ref', newRef);
                    xml = xml.substring(0, afMatch.index) + newAf +
                        xml.substring(afMatch.index + afMatch[0].length);
                }
                if (table.staged) {
                    fs.writeFileSync(this.stagedParts.get(table.path)!, xml, 'utf8');
                } else {
                    this.zip.updateFile(table.path, Buffer.from(xml, 'utf8'));
                }
            } catch {
                // Best-effort per table.
            }
        }
    }
}

export default XlsxUpdater;

/** Explicit name for updating macro-enabled XLSM packages. */
export { XlsxUpdater as XlsmUpdater };
