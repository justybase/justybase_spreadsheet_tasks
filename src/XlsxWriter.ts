import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import archiver from 'archiver';
import { Readable } from 'stream';
import { BigBuffer } from './BigBuffer';
import { CellValue, CellStyle } from './Formats';
import { TextByteSink } from './writers/core/ByteSink';
import { StreamingSheetState } from './StreamingSheetState';
import {
    sanitizeSheetName,
    initColWidths,
    applyHeaderWidths,
    updateColWidthsFromRows,
    assertColumnWidth,
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
    timeToOaDate,
    AUTOFIT_SAMPLE_ROWS,
    resolveAutofitRows,
    assertCompressionLevel,
    DEFAULT_COMPRESSION_LEVEL,
} from './writerHelpers';
import { columnIndexToLetter } from './xmlUtils';
import {
    buildXlsxContentTypes,
    buildXlsxComments,
    buildXlsxCommentsVml,
    buildXlsxRels,
    buildXlsxSharedStrings,
    buildXlsxSheetRels,
    buildXlsxWorkbook,
    CommentSpec,
    ConditionalFormatSpec,
    DataValidationSpec,
    HyperlinkSpec,
    SharedStringValue,
    writeXlsxCell,
    writeXlsxInlineStringCell,
    writeXlsxSharedStringCell,
    writeXlsxSheetHead,
    writeXlsxSheetTail,
    XLSX_REL_COMMENTS,
    XLSX_REL_HYPERLINK,
    XLSX_REL_VML,
    XlsxSheetFeatures,
    XlsxStyleRegistry,
} from './writers/core/xlsxParts';

interface SheetInfo {
    name: string;
    pathInArchive: string;
    hidden: boolean;
    nameInArchive: string;
    sheetId: number;
    rId: string;
    filterHeaderRange: string | null;
    merges?: string[];
    dataValidations?: DataValidationSpec[];
    conditionalFormats?: ConditionalFormatSpec[];
    comments?: CommentSpec[];
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string }>;
    legacyDrawingRId?: string;
    sheetRelsXml?: string;
    extraParts?: Array<{ name: string; data: string }>;
}

export interface XlsxSheetOptions extends XlsxSheetFeatures {
    hidden?: boolean;
    doAutofilter?: boolean;
    sampleRows?: unknown[][];
    /** Conservative autofit from headers and the first rows (default true). */
    autofit?: boolean;
    /** Rows measured by autofit (default 200); 0 measures only headers/samples. */
    autofitRows?: number;
    /** Fixed widths in Excel character units; skips autofit when provided. */
    columnWidths?: Array<number | null | undefined>;
}

/**
 * XLSX (Office Open XML) writer with streaming ZIP output.
 */
export class XlsxWriter {
    private output: fs.WriteStream;
    private archive: archiver.Archiver;
    private readonly _targetPath: string;
    private readonly _tempPath: string;
    private _outputError: Error | null = null;
    private _finalized: boolean = false;

    private sheetCount: number = 0;
    private sheetList: SheetInfo[] = [];
    private sstArray: SharedStringValue[] = [];
    private sstMap: Map<string, number> = new Map();
    private readonly _styles = new XlsxStyleRegistry();
    private sstCntAll: number = 0;
    private _autofilterIsOn: boolean = false;

    private _oaEpoch: number;

    private readonly _useSharedStrings: boolean;

    private readonly stream = new StreamingSheetState();
    private currentColLetters: string[] = [];
    private _streamPending: BigBuffer | null = null;
    private _streamHeadFlushed: boolean = true;
    private _streamHasHeaders: boolean = false;
    private _streamDataRows: number = 0;
    private _streamAutofitRows: number = AUTOFIT_SAMPLE_ROWS;

    /**
     * @param filePath Destination `.xlsx` path on disk.
     */
    /**
     * @param filePath Destination `.xlsx` path on disk.
     * @param options.compressionLevel Deflate level 0-9 (default 3).
     */
    constructor(filePath: string, options: { compressionLevel?: number; useSharedStrings?: boolean } = {}) {
        const compressionLevel = assertCompressionLevel(options.compressionLevel ?? DEFAULT_COMPRESSION_LEVEL);
        this._useSharedStrings = options.useSharedStrings !== false;
        this._targetPath = path.resolve(filePath);
        fs.mkdirSync(path.dirname(this._targetPath), { recursive: true });
        this._tempPath = path.join(
            path.dirname(this._targetPath),
            `.${path.basename(this._targetPath)}.${randomUUID()}.tmp`
        );
        this.output = fs.createWriteStream(this._tempPath);
        this.output.on('error', (err: Error) => { this._outputError = err; });
        this.archive = archiver('zip', { zlib: { level: compressionLevel } });
        this.archive.on('error', (err: Error) => { this._outputError = err; });

        this.archive.pipe(this.output);

        this._oaEpoch = Date.UTC(1899, 11, 30);
    }

    private _getColumnLetter(colIndex: number): string {
        return columnIndexToLetter(colIndex);
    }

    /** @deprecated Internal; kept for unit-test access. Prefer sanitizeSheetName(). */
    private _sanitizeSheetName(name: string): string {
        return sanitizeSheetName(name, this.sheetCount);
    }

    /**
     * Register a worksheet (metadata only). Prefer `writeSheet` / `startSheet` for data.
     */
    addSheet(sheetName: string, hidden: boolean = false): void {
        const sanitizedName = sanitizeSheetName(sheetName, this.sheetCount);

        this.sheetCount++;
        const rId = `rId${this.sheetCount}`;
        const sheetFileName = `sheet${this.sheetCount}.xml`;
        this.sheetList.push({
            name: sanitizedName,
            pathInArchive: `xl/worksheets/${sheetFileName}`,
            hidden: hidden,
            nameInArchive: sheetFileName,
            sheetId: this.sheetCount,
            rId: rId,
            filterHeaderRange: null
        });
    }

    /**
     * Resolve column widths with the conservative autofit: header text and up
     * to 200 sample rows, calibrated character costs plus padding, capped at
     * 200. Fixed `columnWidths` win over autofit.
     */
    private _resolveColumnWidths(
        columnCount: number,
        headers: string[] | null | undefined,
        rows: unknown[][] | undefined,
        options: XlsxSheetOptions,
        doAutofilter: boolean,
    ): number[] {
        const widths = initColWidths(columnCount);
        if (options.columnWidths) {
            for (let i = 0; i < columnCount; i++) {
                const width = options.columnWidths[i];
                if (width !== null && width !== undefined) widths[i] = assertColumnWidth(width, i);
            }
            return widths;
        }
        if (options.autofit === false) return widths;
        if (headers) {
            applyHeaderWidths(widths, headers, columnCount, {
                bold: true,
                filterButton: doAutofilter,
            });
        }
        const samples = options.sampleRows ?? rows;
        const limit = options.sampleRows
            ? samples?.length ?? 0
            : resolveAutofitRows(options.autofitRows);
        if (samples && limit > 0) updateColWidthsFromRows(widths, samples, limit);
        return widths;
    }

    /** Store sheet features and build the worksheet rels/comments parts. */
    private _applySheetFeatures(sheet: SheetInfo, features: XlsxSheetFeatures): void {
        if (features.merges) sheet.merges = features.merges;
        if (features.dataValidations) sheet.dataValidations = features.dataValidations;
        if (features.conditionalFormats) sheet.conditionalFormats = features.conditionalFormats;

        const rels: Array<{ id: string; type: string; target: string; external?: boolean }> = [];
        let nextRelId = 1;
        const hyperlinks: Array<{ spec: HyperlinkSpec; rId?: string }> = [];
        for (const spec of features.hyperlinks ?? []) {
            if (spec.target) {
                const id = `rId${nextRelId++}`;
                rels.push({ id, type: XLSX_REL_HYPERLINK, target: spec.target, external: true });
                hyperlinks.push({ spec, rId: id });
            } else {
                hyperlinks.push({ spec });
            }
        }
        if (hyperlinks.length > 0) sheet.hyperlinks = hyperlinks;

        const comments = features.comments ?? [];
        if (comments.length > 0) {
            const commentsPath = `xl/comments/sheet${sheet.sheetId}.xml`;
            const vmlPath = `xl/drawings/commentsDrawing${sheet.sheetId}.vml`;
            const commentsRid = `rId${nextRelId++}`;
            const vmlRid = `rId${nextRelId++}`;
            rels.push({
                id: commentsRid,
                type: XLSX_REL_COMMENTS,
                target: `../comments/sheet${sheet.sheetId}.xml`,
            });
            rels.push({
                id: vmlRid,
                type: XLSX_REL_VML,
                target: `../drawings/commentsDrawing${sheet.sheetId}.vml`,
            });
            sheet.comments = comments;
            sheet.legacyDrawingRId = vmlRid;
            sheet.extraParts = [
                { name: commentsPath, data: buildXlsxComments(comments) },
                { name: vmlPath, data: buildXlsxCommentsVml(comments) },
            ];
        }

        if (rels.length > 0) sheet.sheetRelsXml = buildXlsxSheetRels(rels);
    }

    private _appendExtraParts(sheet: SheetInfo): void {
        if (sheet.sheetRelsXml) {
            this.archive.append(sheet.sheetRelsXml, {
                name: `xl/worksheets/_rels/${sheet.nameInArchive}.rels`,
            });
        }
        for (const part of sheet.extraParts ?? []) {
            this.archive.append(part.data, { name: part.name });
        }
    }


    startSheet(
        sheetName: string,
        columnCount: number,
        headers?: string[],
        options: XlsxSheetOptions = {}
    ): void {
        assertColumnCount(columnCount);
        if (headers !== undefined) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }

        const { hidden = false, doAutofilter = true } = options;

        this.addSheet(sheetName, hidden);
        const sheet = this.sheetList[this.sheetCount - 1];
        this._applySheetFeatures(sheet, options);

        this.stream.begin(columnCount, doAutofilter && headers !== undefined, new BigBuffer());

        const colLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) {
            colLetters[i] = this._getColumnLetter(i);
        }
        this.currentColLetters = colLetters;

        this.stream.colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            undefined,
            options,
            doAutofilter && headers !== undefined,
        );

        // Widths come from the header and the first AUTOFIT_SAMPLE_ROWS data
        // rows, so the sheet head is delayed until they are measured (xlspy
        // behaves the same way). Until then rows are buffered here.
        this._streamHasHeaders = headers !== undefined;
        this._streamDataRows = 0;
        this._streamAutofitRows = resolveAutofitRows(options.autofitRows);
        this._streamHeadFlushed = false;
        this._streamPending = new BigBuffer();

        if (headers) {
            this.stream.rowNum++;
            const pending = this._streamPending;
            pending.writeString(`<row r="${this.stream.rowNum}">`);
            // Headers are contiguous from column A, so no cell needs `r`.
            for (let c = 0; c < headers.length; c++) {
                this._writeStringCell(pending, headers[c], '', 3);
            }
            pending.writeString('</row>');
        }
        if (this._streamAutofitRows === 0) this._flushStreamHead();
    }

    /** Write the delayed sheet head and everything buffered before it. */
    private _flushStreamHead(): void {
        if (this._streamHeadFlushed) return;
        const bigBuf = this.stream.assertStreaming();
        const pending = this._streamPending;
        writeXlsxSheetHead(bigBuf, {
            dimensionRef: 'A1',
            columnCount: this.stream.endCol,
            isFirstSheet: this.sheetCount === 1,
            hasHeaders: this._streamHasHeaders,
            colWidths: this.stream.colWidths,
        });
        if (pending) {
            for (const chunk of pending.getChunks()) bigBuf.write(chunk);
        }
        this._streamPending = null;
        this._streamHeadFlushed = true;
    }

    private _writeStreamRow(bigBuf: BigBuffer, row: CellValue[]): void {
        this.stream.rowNum++;
        assertRowCount(this.stream.rowNum);
        bigBuf.writeString(`<row r="${this.stream.rowNum}">`);

        let prevCol = -1;
        for (let c = 0; c < row.length; c++) {
            prevCol = writeXlsxCell(
                bigBuf,
                this,
                row[c],
                c,
                prevCol,
                this.currentColLetters[c],
                this.stream.rowNum,
            );
        }
        bigBuf.writeString('</row>');
    }

    /** Write one data row in streaming mode. */
    writeRow(row: CellValue[]): void {
        const bigBuf = this.stream.assertStreaming();

        if (row.length !== this.stream.endCol - this.stream.startCol) {
            throw new Error(
                `Row length mismatch. Expected ${this.stream.endCol - this.stream.startCol} columns, got ${row.length}`
            );
        }

        if (!this._streamHeadFlushed && this._streamPending) {
            this._streamDataRows++;
            updateColWidthsFromRows(this.stream.colWidths, [row]);
            this._writeStreamRow(this._streamPending, row);
            if (this._streamDataRows >= this._streamAutofitRows) this._flushStreamHead();
            return;
        }
        this._writeStreamRow(bigBuf, row);
    }

    /** Finish the current streaming sheet and append it to the archive. */
    endSheet(): void {
        this._flushStreamHead();
        const bigBuf = this.stream.assertStreaming();
        const sheet = this.sheetList[this.sheetCount - 1];
        const filterRange = writeXlsxSheetTail(bigBuf, {
            doAutofilter: this.stream.doAutofilter,
            columnCount: this.stream.endCol,
            lastColumnLetter: this.currentColLetters[this.stream.endCol - 1] ?? 'A',
            rowNum: this.stream.rowNum,
            sheetName: sheet.name,
            merges: sheet.merges,
            hyperlinks: sheet.hyperlinks,
            dataValidations: sheet.dataValidations,
            conditionalFormats: sheet.conditionalFormats,
            legacyDrawingRId: sheet.legacyDrawingRId,
            registerDxf: style => this._styles.registerDxf(style),
        });
        if (filterRange !== null) {
            this._autofilterIsOn = true;
            sheet.filterHeaderRange = filterRange;
        }

        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: sheet.pathInArchive
        });
        this._appendExtraParts(sheet);

        this.stream.end();
        this.currentColLetters = [];
        this._streamPending = null;
        this._streamHeadFlushed = true;
        this._streamHasHeaders = false;
        this._streamDataRows = 0;
        this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    }

    /**
     * Write an entire sheet in one call (batch mode).
     *
     * The third argument keeps the legacy `doAutofilter` boolean; pass
     * {@link XlsxSheetOptions} instead to add features (merges, hyperlinks,
     * validations, conditional formats, comments) and width control.
     */
    writeSheet(
        rows: CellValue[][],
        headers: string[] | null = null,
        doAutofilterOrOptions: boolean | XlsxSheetOptions = true,
        extraOptions: XlsxSheetOptions = {},
    ): void {
        let doAutofilter = true;
        let options: XlsxSheetOptions;
        if (typeof doAutofilterOrOptions === 'boolean') {
            // The options object is the newer API; an explicit `doAutofilter`
            // there wins over the legacy positional boolean.
            doAutofilter = extraOptions.doAutofilter ?? doAutofilterOrOptions;
            options = { ...extraOptions, doAutofilter };
        } else {
            options = { ...doAutofilterOrOptions, ...extraOptions };
            doAutofilter = options.doAutofilter !== false;
        }

        const bigBuf = new BigBuffer();

        let columnCount = 0;
        if (rows.length > 0) {
            columnCount = rows[0].length;
        } else if (headers) {
            columnCount = headers.length;
        }

        if (columnCount > 0) assertColumnCount(columnCount);
        const totalRows = rows.length + (headers ? 1 : 0);
        assertRowCount(totalRows);
        if (headers) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        for (let r = 0; r < rows.length; r++) {
            assertRowFitsColumns('row', rows[r].length, columnCount);
        }

        const colLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) {
            colLetters[i] = this._getColumnLetter(i);
        }

        const colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            rows,
            options,
            doAutofilter && headers !== null && headers !== undefined,
        );

        const dimensionRef = totalRows > 0 && columnCount > 0
            ? `A1:${colLetters[columnCount - 1]}${totalRows}`
            : 'A1';
        writeXlsxSheetHead(bigBuf, {
            dimensionRef,
            columnCount,
            isFirstSheet: this.sheetCount === 1,
            hasHeaders: headers !== null && headers !== undefined,
            colWidths,
        });

        let rowNum = 0;

        if (headers) {
            rowNum++;
            bigBuf.writeString(`<row r="${rowNum}">`);
            // Headers are contiguous from column A, so no cell needs `r`.
            for (let c = 0; c < headers.length; c++) {
                this._writeStringCell(bigBuf, headers[c], '', 3);
            }
            bigBuf.writeString('</row>');
        }

        for (let r = 0; r < rows.length; r++) {
            rowNum++;
            bigBuf.writeString(`<row r="${rowNum}">`);
            const row = rows[r];
            let prevCol = -1;
            for (let c = 0; c < row.length; c++) {
                prevCol = writeXlsxCell(bigBuf, this, row[c], c, prevCol, colLetters[c], rowNum);
            }
            bigBuf.writeString('</row>');
        }

        const sheet = this.sheetList[this.sheetCount - 1];
        if (!sheet) {
            throw new Error('Call addSheet() before writeSheet()');
        }
        this._applySheetFeatures(sheet, options);

        const filterRange = writeXlsxSheetTail(bigBuf, {
            doAutofilter: doAutofilter && headers !== null && headers !== undefined,
            columnCount,
            lastColumnLetter: columnCount > 0 ? colLetters[columnCount - 1] : 'A',
            rowNum: totalRows,
            sheetName: sheet.name,
            merges: sheet.merges,
            hyperlinks: sheet.hyperlinks,
            dataValidations: sheet.dataValidations,
            conditionalFormats: sheet.conditionalFormats,
            legacyDrawingRId: sheet.legacyDrawingRId,
            registerDxf: style => this._styles.registerDxf(style),
        });
        if (filterRange !== null) {
            this._autofilterIsOn = true;
            sheet.filterHeaderRange = filterRange;
        }

        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: sheet.pathInArchive
        });
        this._appendExtraParts(sheet);
    }

    /** `ref` is the cell's `r` attribute from `cellRefAttr`, or '' when omitted. */
    private _writeStringCell(bigBuf: TextByteSink, val: SharedStringValue, ref: string, styleOverride?: number): void {
        assertCellTextLength(typeof val === 'string' ? val : val.toString());
        if (this._useSharedStrings) {
            writeXlsxSharedStringCell(bigBuf, { values: this.sstArray, index: this.sstMap }, val, ref, styleOverride);
            this.sstCntAll++;
            return;
        }
        writeXlsxInlineStringCell(bigBuf, val, ref, styleOverride);
    }

    // XlsxCellContext implementation used by the shared cell writer.
    writeStringCell(bigBuf: TextByteSink, val: SharedStringValue, ref: string, styleOverride?: number): void {
        this._writeStringCell(bigBuf, val, ref, styleOverride);
    }

    registerStyle(style: CellStyle): number {
        return this._styles.register(style);
    }

    registerFormat(fmtString: string): number {
        return this._styles.registerFormat(fmtString);
    }

    toOaDate(date: Date): number {
        return this._toOADate(date);
    }

    private _toOADate(date: Date): number {
        // Use UTC time to avoid timezone conversion issues; handles the
        // fictitious 1900-02-29 (dates before 1900-03-01 serialize one lower).
        return timeToOaDate(date.getTime(), this._oaEpoch);
    }

    /** Finalize the ZIP package and close the output stream. */
    finalize(): Promise<void> {
        return new Promise((resolve, reject) => {
            if (this._finalized) {
                reject(new Error('XlsxWriter: already finalized'));
                return;
            }
            this._finalized = true;
            let settled = false;
            const cleanupTemp = (): void => {
                try {
                    if (fs.existsSync(this._tempPath)) fs.rmSync(this._tempPath);
                } catch {
                    // Preserve original error.
                }
            };
            const fail = (err: unknown): void => {
                if (settled) return;
                settled = true;
                try { this.archive.abort(); } catch { /* preserve */ }
                try { this.output.destroy(); } catch { /* preserve */ }
                cleanupTemp();
                reject(err instanceof Error ? err : new Error(String(err)));
            };
            const succeed = (): void => {
                if (settled) return;
                if (this._outputError) {
                    fail(this._outputError);
                    return;
                }
                settled = true;
                try {
                    // Ensure data hits disk before publishing (xlspy parity, no fsync of dir).
                    const fd = fs.openSync(this._tempPath, 'r+');
                    try { fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
                    // Preserve the mode of a file being replaced (the temp file
                    // was created with the process umask).
                    try {
                        const existing = fs.statSync(this._targetPath);
                        fs.chmodSync(this._tempPath, existing.mode & 0o777);
                    } catch {
                        // Target does not exist yet.
                    }
                    fs.renameSync(this._tempPath, this._targetPath);
                    resolve();
                } catch (err) {
                    cleanupTemp();
                    reject(err instanceof Error ? err : new Error(String(err)));
                }
            };
            try {
                if (this._outputError) throw this._outputError;
                this._writeSharedStrings();
                this._writeStyles();
                this._writeWorkbook();
                this._writeContentTypes();
                this._writeRels();

                this.output.once('close', succeed);
                this.output.once('error', fail);
                this.archive.once('error', fail);

                this.archive.finalize();
            } catch (err) {
                fail(err);
            }
        });
    }

    private _writeSharedStrings(): void {
        this.archive.append(
            buildXlsxSharedStrings(this.sstArray, this.sstCntAll),
            { name: 'xl/sharedStrings.xml' },
        );
    }

    private _writeStyles(): void {
        this.archive.append(this._styles.toXml(), { name: 'xl/styles.xml' });
    }

    private _writeWorkbook(): void {
        this.archive.append(
            buildXlsxWorkbook(this.sheetList, this._autofilterIsOn),
            { name: 'xl/workbook.xml' },
        );
    }

    private _writeContentTypes(): void {
        const commentParts = this.sheetList
            .filter(sheet => (sheet.comments?.length ?? 0) > 0)
            .map(sheet => ({ path: `xl/comments/sheet${sheet.sheetId}.xml` }));
        this.archive.append(
            buildXlsxContentTypes(this.sheetList, commentParts),
            { name: '[Content_Types].xml' },
        );
    }

    private _writeRels(): void {
        const rels = buildXlsxRels(this.sheetList);
        this.archive.append(rels.global, { name: '_rels/.rels' });
        this.archive.append(rels.workbook, { name: 'xl/_rels/workbook.xml.rels' });
    }
}

export default XlsxWriter;
