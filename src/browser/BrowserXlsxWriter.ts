/**
 * BrowserXlsxWriter - XLSX writer for browsers.
 *
 * The package parts and the cell XML are produced by the shared writer core
 * (`src/writers/core`), the same code the Node `XlsxWriter` uses. Only the ZIP
 * sink differs (BrowserZip instead of archiver).
 */
import { BrowserBigBuffer } from './BrowserBigBuffer';
import { BrowserZip } from './BrowserZip';
import { CellValue, CellStyle } from '../Formats';
import { columnIndexToLetter } from '../xmlUtils';
import {
    applyHeaderWidths,
    assertCellTextLength as assertBrowserCellText,
    assertColumnCount as assertBrowserColumnCount,
    assertRowFitsColumns as assertBrowserRowFits,
    assertRowCount as assertBrowserRowCount,
    initColWidths,
    timeToOaDate as timeToOaDateBrowser,
    updateColWidthsFromRows,
    assertColumnWidth,
    AUTOFIT_SAMPLE_ROWS,
    resolveAutofitRows,
} from '../writerHelpers';
import {
    buildXlsxComments,
    buildXlsxCommentsVml,
    buildXlsxContentTypes,
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
    XlsxSheetEntry,
    XlsxSheetFeatures,
    XlsxStyleRegistry,
} from '../writers/core/xlsxParts';

export type BrowserHeaderStyle = 'bold' | 'fill' | 'bold+fill';

interface BrowserSheetInfo extends XlsxSheetEntry {
    merges?: string[];
    dataValidations?: DataValidationSpec[];
    conditionalFormats?: ConditionalFormatSpec[];
    comments?: CommentSpec[];
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string }>;
    legacyDrawingRId?: string;
    sheetRelsXml?: string;
    extraParts?: Array<{ name: string; data: string }>;
}

export interface BrowserSheetOptions extends XlsxSheetFeatures {
    doAutofilter?: boolean;
    headerStyle?: BrowserHeaderStyle;
    /** Conservative autofit from headers and the first rows (default true). */
    autofit?: boolean;
    /** Rows measured by autofit (default 200); 0 measures only headers/samples. */
    autofitRows?: number;
    /** Fixed widths in Excel character units; skips autofit when provided. */
    columnWidths?: Array<number | null | undefined>;
}

export interface BrowserStartSheetOptions extends BrowserSheetOptions {
    hidden?: boolean;
    sampleRows?: unknown[][];
}

export interface BrowserXlsxWriterOptions {
    useSharedStrings?: boolean;
}

function headerStyleIndex(style: BrowserHeaderStyle): number {
    if (style === 'fill') return 4;
    if (style === 'bold+fill') return 5;
    return 3;
}

export class BrowserXlsxWriter {
    private _zip: BrowserZip;
    private _sheetCount = 0;
    private _sheets: BrowserSheetInfo[] = [];
    private _sstValues: SharedStringValue[] = [];
    private _sstIndex = new Map<string, number>();
    private _sstCount = 0;
    private _autofilterOn = false;
    private _oaEpoch = Date.UTC(1899, 11, 30);
    private _useSharedStrings: boolean;

    private readonly _styles = new XlsxStyleRegistry();

    private _streamBuffer: BrowserBigBuffer | null = null;
    private _streamRowNum = 0;
    private _streamColumnCount = 0;
    private _streamDoAutofilter = false;
    private _streamColLetters: string[] = [];
    private _streamColWidths: number[] = [];
    private _streamPending: BrowserBigBuffer | null = null;
    private _streamHeadFlushed = true;
    private _streamHasHeaders = false;
    private _streamDataRows = 0;
    private _streamAutofitRows = AUTOFIT_SAMPLE_ROWS;

    constructor(options: BrowserXlsxWriterOptions = {}) {
        this._zip = new BrowserZip();
        this._useSharedStrings = options.useSharedStrings !== false;
    }

    addSheet(name: string, hidden = false): void {
        const sheetName = this._sanitizeSheetName(name);
        this._sheetCount++;
        const fileName = `sheet${this._sheetCount}.xml`;
        this._sheets.push({
            name: sheetName,
            pathInArchive: `xl/worksheets/${fileName}`,
            nameInArchive: fileName,
            hidden,
            sheetId: this._sheetCount,
            rId: `rId${this._sheetCount}`,
            filterHeaderRange: null,
        });
    }

    /** @deprecated Internal; kept for unit-test access. */
    _cl(index: number): string {
        return columnIndexToLetter(index);
    }

    private _sanitizeSheetName(name: string): string {
        if (!name || typeof name !== 'string') return `Sheet${this._sheetCount + 1}`;
        let value = name.replace(/[\\/*?[\]:]/g, '_');
        if (value.length > 31) value = value.substring(0, 31);
        if (!value.trim()) value = `Sheet${this._sheetCount + 1}`;
        return value;
    }

    // XlsxCellContext ---------------------------------------------------

    writeStringCell(sink: BrowserBigBuffer, value: SharedStringValue, ref: string, styleOverride?: number): void {
        assertBrowserCellText(typeof value === 'string' ? value : value.toString());
        if (this._useSharedStrings) {
            writeXlsxSharedStringCell(
                sink,
                { values: this._sstValues, index: this._sstIndex },
                value,
                ref,
                styleOverride,
            );
            this._sstCount++;
            return;
        }
        writeXlsxInlineStringCell(sink, value, ref, styleOverride);
    }

    registerStyle(style: CellStyle): number {
        return this._styles.register(style);
    }

    registerFormat(format: string): number {
        return this._styles.registerFormat(format);
    }

    toOaDate(date: Date): number {
        return timeToOaDateBrowser(date.getTime(), this._oaEpoch);
    }

    // Sheet features/widths ---------------------------------------------

    /**
     * Resolve column widths: explicit `columnWidths` win, otherwise plain
     * autofit from headers and up to 200 sample rows, capped at 200.
     */
    private _resolveColumnWidths(
        columnCount: number,
        headers: string[] | null | undefined,
        rows: unknown[][] | undefined,
        options: BrowserSheetOptions & { sampleRows?: unknown[][] },
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
                bold: options.headerStyle !== 'fill',
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
    private _applySheetFeatures(sheet: BrowserSheetInfo, features: BrowserSheetOptions): void {
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

    private _appendExtraParts(sheet: BrowserSheetInfo): void {
        if (sheet.sheetRelsXml) {
            this._zip.addFile(`xl/worksheets/_rels/${sheet.nameInArchive}.rels`, sheet.sheetRelsXml);
        }
        for (const part of sheet.extraParts ?? []) {
            this._zip.addFile(part.name, part.data);
        }
    }

    // Streaming API -----------------------------------------------------

    startSheet(
        sheetName: string,
        columnCount: number,
        headers?: string[],
        options: BrowserStartSheetOptions = {},
    ): void {
        if (this._streamBuffer !== null) throw new Error('Already streaming');
        assertBrowserColumnCount(columnCount);
        if (headers !== undefined) {
            assertBrowserRowFits('header', headers.length, columnCount);
            for (const header of headers) assertBrowserCellText(header);
        }

        const { hidden = false, doAutofilter = true, headerStyle = 'bold' } = options;
        this.addSheet(sheetName, hidden);
        const sheet = this._sheets[this._sheetCount - 1];
        this._applySheetFeatures(sheet, options);
        this._streamBuffer = new BrowserBigBuffer();
        this._streamRowNum = 0;
        this._streamColumnCount = columnCount;
        this._streamDoAutofilter = doAutofilter && headers !== undefined;

        this._streamColWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            undefined,
            options,
            doAutofilter && headers !== undefined,
        );
        this._streamColLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) this._streamColLetters[i] = columnIndexToLetter(i);

        // Delay the head until the first AUTOFIT_SAMPLE_ROWS rows are measured.
        this._streamHasHeaders = headers !== undefined;
        this._streamDataRows = 0;
        this._streamAutofitRows = resolveAutofitRows(options.autofitRows);
        this._streamHeadFlushed = false;
        this._streamPending = new BrowserBigBuffer();

        if (headers) {
            const pending = this._streamPending;
            this._streamRowNum++;
            pending.writeString(`<row r="${this._streamRowNum}">`);
            const style = headerStyleIndex(headerStyle);
            for (let c = 0; c < headers.length; c++) {
                this.writeStringCell(pending, headers[c], '', style);
            }
            pending.writeString('</row>');
        }
        if (this._streamAutofitRows === 0) this._flushStreamHead();
    }

    /** Write the delayed sheet head and everything buffered before it. */
    private _flushStreamHead(): void {
        if (this._streamHeadFlushed) return;
        if (this._streamBuffer === null) throw new Error('Not streaming');
        const buffer = this._streamBuffer;
        const pending = this._streamPending;
        writeXlsxSheetHead(buffer, {
            dimensionRef: 'A1',
            columnCount: this._streamColumnCount,
            isFirstSheet: this._sheetCount === 1,
            hasHeaders: this._streamHasHeaders,
            colWidths: this._streamColWidths,
        });
        if (pending) {
            for (const chunk of pending.getChunks()) buffer.write(chunk);
        }
        this._streamPending = null;
        this._streamHeadFlushed = true;
    }

    private _writeStreamRow(buffer: BrowserBigBuffer, row: CellValue[]): void {
        this._streamRowNum++;
        assertBrowserRowCount(this._streamRowNum);
        buffer.writeString(`<row r="${this._streamRowNum}">`);
        let prevCol = -1;
        for (let c = 0; c < row.length; c++) {
            prevCol = writeXlsxCell(
                buffer,
                this,
                row[c],
                c,
                prevCol,
                this._streamColLetters[c],
                this._streamRowNum,
            );
        }
        buffer.writeString('</row>');
    }

    writeRow(row: CellValue[]): void {
        if (this._streamBuffer === null) throw new Error('Not streaming');
        if (row.length !== this._streamColumnCount) {
            throw new Error(`Row length mismatch. Expected ${this._streamColumnCount} columns, got ${row.length}`);
        }
        if (!this._streamHeadFlushed && this._streamPending) {
            this._streamDataRows++;
            updateColWidthsFromRows(this._streamColWidths, [row]);
            this._writeStreamRow(this._streamPending, row);
            if (this._streamDataRows >= this._streamAutofitRows) this._flushStreamHead();
            return;
        }
        this._writeStreamRow(this._streamBuffer, row);
    }

    endSheet(): void {
        this._flushStreamHead();
        if (this._streamBuffer === null) throw new Error('Not streaming');
        const buffer = this._streamBuffer;
        const sheet = this._sheets[this._sheetCount - 1];
        const filterRange = writeXlsxSheetTail(buffer, {
            doAutofilter: this._streamDoAutofilter,
            columnCount: this._streamColumnCount,
            lastColumnLetter: this._streamColLetters[this._streamColumnCount - 1] ?? 'A',
            rowNum: this._streamRowNum,
            sheetName: sheet.name,
            merges: sheet.merges,
            hyperlinks: sheet.hyperlinks,
            dataValidations: sheet.dataValidations,
            conditionalFormats: sheet.conditionalFormats,
            legacyDrawingRId: sheet.legacyDrawingRId,
            registerDxf: style => this._styles.registerDxf(style),
        });
        if (filterRange !== null) {
            this._autofilterOn = true;
            sheet.filterHeaderRange = filterRange;
        }
        this._zip.addFile(sheet.pathInArchive, buffer.toUint8Array());
        this._appendExtraParts(sheet);
        this._streamBuffer = null;
        this._streamPending = null;
        this._streamHeadFlushed = true;
        this._streamHasHeaders = false;
        this._streamDataRows = 0;
        this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
        this._streamColWidths = [];
    }

    // Batch API ---------------------------------------------------------

    writeSheet(
        rows: CellValue[][],
        headers: string[] | null = null,
        options: BrowserSheetOptions = {},
    ): void {
        const doAutofilter = options.doAutofilter !== false;
        const headerStyle = options.headerStyle ?? 'bold';
        const buffer = new BrowserBigBuffer();

        const columnCount = rows.length > 0 ? rows[0].length : (headers ? headers.length : 0);
        if (columnCount > 0) assertBrowserColumnCount(columnCount);
        const totalRows = rows.length + (headers ? 1 : 0);
        assertBrowserRowCount(totalRows);
        if (headers) {
            assertBrowserRowFits('header', headers.length, columnCount);
            for (const header of headers) assertBrowserCellText(header);
        }
        for (const row of rows) assertBrowserRowFits('row', row.length, columnCount);
        const sheet = this._sheets[this._sheetCount - 1];
        if (!sheet) throw new Error('Call addSheet() before writeSheet()');

        const colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            rows,
            options,
            doAutofilter && headers !== null,
        );

        const colLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) colLetters[i] = columnIndexToLetter(i);

        const dimensionRef = totalRows > 0 && columnCount > 0
            ? `A1:${colLetters[columnCount - 1]}${totalRows}`
            : 'A1';
        writeXlsxSheetHead(buffer, {
            dimensionRef,
            columnCount,
            isFirstSheet: this._sheetCount === 1,
            hasHeaders: headers !== null,
            colWidths,
        });

        let rowNum = 0;
        if (headers) {
            rowNum++;
            buffer.writeString(`<row r="${rowNum}">`);
            const style = headerStyleIndex(headerStyle);
            for (let c = 0; c < headers.length; c++) {
                this.writeStringCell(buffer, headers[c], '', style);
            }
            buffer.writeString('</row>');
        }

        for (const row of rows) {
            rowNum++;
            buffer.writeString(`<row r="${rowNum}">`);
            let prevCol = -1;
            for (let c = 0; c < row.length; c++) {
                prevCol = writeXlsxCell(buffer, this, row[c], c, prevCol, colLetters[c], rowNum);
            }
            buffer.writeString('</row>');
        }

        this._applySheetFeatures(sheet, options);
        const filterRange = writeXlsxSheetTail(buffer, {
            doAutofilter: doAutofilter && headers !== null,
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
            this._autofilterOn = true;
            sheet.filterHeaderRange = filterRange;
        }

        this._zip.addFile(sheet.pathInArchive, buffer.toUint8Array());
        this._appendExtraParts(sheet);
    }

    // Package -----------------------------------------------------------

    finalize(): Blob {
        this._zip.addFile(
            'xl/sharedStrings.xml',
            buildXlsxSharedStrings(this._sstValues, this._sstCount),
        );
        this._zip.addFile('xl/styles.xml', this._styles.toXml());
        this._zip.addFile('xl/workbook.xml', buildXlsxWorkbook(this._sheets, this._autofilterOn));
        const commentParts = this._sheets
            .filter(sheet => (sheet.comments?.length ?? 0) > 0)
            .map(sheet => ({ path: `xl/comments/sheet${sheet.sheetId}.xml` }));
        this._zip.addFile('[Content_Types].xml', buildXlsxContentTypes(this._sheets, commentParts));
        const rels = buildXlsxRels(this._sheets);
        this._zip.addFile('_rels/.rels', rels.global);
        this._zip.addFile('xl/_rels/workbook.xml.rels', rels.workbook);
        return this._zip.toBlob();
    }
}

export default BrowserXlsxWriter;
