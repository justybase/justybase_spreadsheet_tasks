/**
 * BrowserXlsxWriter - XLSX writer for browsers.
 *
 * The package parts and the cell XML are produced by the shared writer core
 * (`src/writers/core`), the same code the Node `XlsxWriter` uses. Only the ZIP
 * sink differs (BrowserZip instead of archiver).
 */
import { BrowserBigBuffer } from './BrowserBigBuffer';
import { BrowserZip } from './BrowserZip';
import { CellValue, CellStyle, mergeCellStyles } from '../Formats';
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
    SheetLayoutOptions,
    buildXlsxTableXml,
    writeXlsxCell,
    writeXlsxInlineStringCell,
    writeXlsxSharedStringCell,
    writeXlsxSheetHead,
    writeXlsxSheetTail,
    XLSX_REL_COMMENTS,
    XLSX_REL_HYPERLINK,
    XLSX_REL_TABLE,
    XLSX_REL_VML,
    XlsxSheetEntry,
    XlsxSheetFeatures,
    XlsxStyleRegistry,
    xlsxRowAttributes,
} from '../writers/core/xlsxParts';
import {
    FreezePane,
    NormalizedTableSpec,
    Outline,
    SheetProtectionSpec,
    StyleRangeSet,
    buildOutline,
    normalizeGroups,
    normalizeTableSpec,
    parseRangeRef,
    resolveFreezePanes,
    tableTotalsOverrides,
    TableTotalsOverride,
    validateRowHeight,
    validateZoom,
} from '../writers/core/sheetLayout';

/** Apply table totals overrides to one row, preserving the caller's style/cache. */
function withTableTotals(row: CellValue[], overrides: TableTotalsOverride[]): CellValue[] {
    let result: CellValue[] | null = null;
    for (const override of overrides) {
        if (override.col < 0) continue;
        const original = row[override.col];
        const record = original !== null && typeof original === 'object' && !(original instanceof Date)
            ? original as unknown as Record<string, unknown>
            : null;
        const style = record && record.style ? record.style as CellStyle : undefined;
        const cachedValue = record?.value;
        const cached = typeof cachedValue === 'string' || typeof cachedValue === 'number'
            || typeof cachedValue === 'boolean' || cachedValue instanceof Date
            ? cachedValue
            : (typeof original === 'number' ? original : undefined);
        const replacement: CellValue = override.text !== undefined
            ? (style ? { value: override.text, style } : override.text)
            : {
                formula: override.formula as string,
                ...(cached !== undefined ? { value: cached } : {}),
                ...(style ? { style } : {}),
            } as CellValue;
        if (!result) result = row.slice();
        result[override.col] = replacement;
    }
    return result ?? row;
}

export type BrowserHeaderStyle = 'bold' | 'fill' | 'bold+fill';

interface ResolvedBrowserLayout {
    headerStyle: CellStyle | null;
    styleRanges: StyleRangeSet;
    rowHeights: Record<number, number>;
    hiddenRows: Set<number>;
    hiddenColumns: number[];
    rowOutline: Outline;
    columnOutline: Outline;
    freezePanes: FreezePane | null;
    showGridLines: boolean;
    showRowColHeaders: boolean;
    showZeroValues: boolean;
    rightToLeft: boolean;
    zoom?: number;
    tabColor?: import('../Formats').StyleColor;
    sheetProtection?: SheetProtectionSpec;
    autoFilterRef?: string;
    tables: NormalizedTableSpec[];
}

interface BrowserSheetInfo extends XlsxSheetEntry {
    merges?: string[];
    dataValidations?: DataValidationSpec[];
    conditionalFormats?: ConditionalFormatSpec[];
    comments?: CommentSpec[];
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string }>;
    legacyDrawingRId?: string;
    sheetRelsXml?: string;
    extraParts?: Array<{ name: string; data: string }>;
    layout?: ResolvedBrowserLayout;
    tableParts?: Array<{ id: number; path: string; xml: string; rId: string }>;
}

export interface BrowserSheetOptions extends XlsxSheetFeatures, Omit<SheetLayoutOptions, 'headerStyle'> {
    doAutofilter?: boolean;
    /** Preset header style; a `headerStyle` CellStyle overrides it. */
    headerPreset?: BrowserHeaderStyle;
    /** Preset name (`bold`, `fill`, `bold+fill`) or an arbitrary cell style. */
    headerStyle?: BrowserHeaderStyle | CellStyle;
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

export class BrowserXlsxWriter {
    private _zip: BrowserZip;
    private _sheetCount = 0;
    private _sheets: BrowserSheetInfo[] = [];
    private _sstValues: SharedStringValue[] = [];
    private _sstIndex = new Map<string, number>();
    private _sstCount = 0;
    private _autofilterOn = false;
    private _nextTableId = 1;
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
    private _streamTableTotals: Map<number, TableTotalsOverride[]> = new Map();

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
        headerStyle?: CellStyle,
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
                bold: headerStyle ? headerStyle.font?.bold === true : true,
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

    private _presetHeaderStyle(options: BrowserSheetOptions): CellStyle | null {
        if (options.headerStyle && typeof options.headerStyle !== 'string') {
            // Match XlsxWriter: the explicit style is layered over bold.
            return mergeCellStyles({ font: { bold: true } }, options.headerStyle);
        }
        const preset = options.headerPreset
            ?? (typeof options.headerStyle === 'string' ? options.headerStyle : undefined)
            ?? 'bold';
        if (preset === 'bold') return null;
        const base: CellStyle = {};
        if (preset === 'bold+fill') base.font = { bold: true };
        if (preset === 'fill' || preset === 'bold+fill') {
            base.fill = { patternType: 'solid', fgColor: 'FFD9E1F2' };
        }
        return base;
    }

    private _headerStyleIndex(layout: ResolvedBrowserLayout): number {
        return layout.headerStyle ? this._styles.register(layout.headerStyle) : 3;
    }

    private _resolveLayout(
        options: BrowserSheetOptions,
        hasHeaders: boolean,
        columnCount: number,
        headers: string[] | null | undefined,
        rowCount: number = 0,
    ): ResolvedBrowserLayout {
        const rowHeights: Record<number, number> = {};
        for (const [key, value] of Object.entries(options.rowHeights ?? {})) {
            const row = Number(key);
            if (!Number.isInteger(row) || row < 0) {
                throw new RangeError(`rowHeights keys must be non-negative row indices, got '${key}'`);
            }
            rowHeights[row] = validateRowHeight(value);
        }

        const hiddenRows = new Set<number>();
        for (const row of options.hiddenRows ?? []) {
            if (!Number.isInteger(row) || row < 0) {
                throw new RangeError(`hiddenRows must contain non-negative row indices, got ${row}`);
            }
            hiddenRows.add(row);
        }

        const hiddenColumns: number[] = [];
        for (const col of options.hiddenColumns ?? []) {
            if (!Number.isInteger(col) || col < 0 || col >= columnCount) {
                throw new RangeError(`hiddenColumns index ${col} is outside the sheet`);
            }
            hiddenColumns.push(col);
        }

        const rowGroups = normalizeGroups(options.rowGroups, 'row');
        const columnGroups = normalizeGroups(options.columnGroups, 'column');
        const rowMax = Math.max(
            rowCount - 1,
            ...rowGroups.map(group => group.end),
            ...Object.keys(rowHeights).map(Number),
            ...hiddenRows,
            0,
        );
        const columnMax = Math.max(
            columnCount - 1,
            ...columnGroups.map(group => group.end),
            ...hiddenColumns,
            0,
        );
        const rowOutline = buildOutline(rowGroups, rowMax);
        const columnOutline = buildOutline(columnGroups, columnMax);

        const tables: NormalizedTableSpec[] = [];
        for (const spec of options.tables ?? []) {
            let columns = spec.columns;
            if ((!columns || columns.length === 0) && headers) {
                const range = parseRangeRef(spec.ref);
                columns = headers.slice(range.c1, range.c2 + 1);
            }
            tables.push(normalizeTableSpec({ ...spec, columns }, this._nextTableId++));
        }

        if (options.autoFilterRef !== undefined) parseRangeRef(options.autoFilterRef);

        return {
            headerStyle: this._presetHeaderStyle(options),
            styleRanges: new StyleRangeSet(options.styles),
            rowHeights,
            hiddenRows,
            hiddenColumns,
            rowOutline,
            columnOutline,
            freezePanes: resolveFreezePanes(options.freezePanes, hasHeaders),
            showGridLines: options.showGridLines !== false,
            showRowColHeaders: options.showRowColHeaders !== false,
            showZeroValues: options.showZeroValues !== false,
            rightToLeft: options.rightToLeft === true,
            zoom: validateZoom(options.zoom),
            tabColor: options.tabColor,
            sheetProtection: options.sheetProtection,
            autoFilterRef: options.autoFilterRef,
            tables,
        };
    }

    private _buildTableParts(sheet: BrowserSheetInfo, layout: ResolvedBrowserLayout, nextRelId: number): number {
        if (layout.tables.length === 0) return nextRelId;
        const parts: Array<{ id: number; path: string; xml: string; rId: string }> = [];
        for (const table of layout.tables) {
            const id = this._nextTableId - layout.tables.length + parts.length;
            const path = `xl/tables/table${id}.xml`;
            parts.push({
                id,
                path,
                xml: buildXlsxTableXml(table, id),
                rId: `rId${nextRelId++}`,
            });
        }
        sheet.tableParts = parts;
        sheet.extraParts = [...(sheet.extraParts ?? []), ...parts.map(part => ({ name: part.path, data: part.xml }))];
        return nextRelId;
    }

    /** Store sheet features and build the worksheet rels/comments parts. */
    private _applySheetFeatures(sheet: BrowserSheetInfo, features: BrowserSheetOptions, layout?: ResolvedBrowserLayout): void {
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

        if (layout && layout.tables.length > 0) {
            nextRelId = this._buildTableParts(sheet, layout, nextRelId);
            for (const part of sheet.tableParts ?? []) {
                rels.push({
                    id: part.rId,
                    type: XLSX_REL_TABLE,
                    target: `../tables/${part.path.split('/').pop()}`,
                });
            }
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

        const { hidden = false, doAutofilter = true } = options;
        this.addSheet(sheetName, hidden);
        const sheet = this._sheets[this._sheetCount - 1];
        const layout = this._resolveLayout(options, headers !== undefined, columnCount, headers);
        sheet.layout = layout;
        this._streamTableTotals = new Map();
        for (const override of tableTotalsOverrides(layout.tables)) {
            const list = this._streamTableTotals.get(override.row) ?? [];
            list.push(override);
            this._streamTableTotals.set(override.row, list);
        }
        this._applySheetFeatures(sheet, options, layout);
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
            layout.headerStyle ?? undefined,
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
            pending.writeString(`<row r="${this._streamRowNum}"${xlsxRowAttributes(0, layout)}>`);
            const headerStyle = this._headerStyleIndex(layout);
            const baseHeader = layout.headerStyle ?? { font: { bold: true } };
            for (let c = 0; c < headers.length; c++) {
                const rangeStyle = layout.styleRanges.styleFor(0, c);
                const style = rangeStyle
                    ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle))
                    : headerStyle;
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
        const layout = this._sheets[this._sheetCount - 1]?.layout;
        writeXlsxSheetHead(buffer, {
            dimensionRef: 'A1',
            columnCount: this._streamColumnCount,
            isFirstSheet: this._sheetCount === 1,
            hasHeaders: this._streamHasHeaders,
            colWidths: this._streamColWidths,
            freezePanes: layout?.freezePanes,
            showGridLines: layout?.showGridLines,
            showRowColHeaders: layout?.showRowColHeaders,
            showZeroValues: layout?.showZeroValues,
            rightToLeft: layout?.rightToLeft,
            zoom: layout?.zoom,
            tabColor: layout?.tabColor,
            hiddenColumns: layout?.hiddenColumns,
            columnOutline: layout?.columnOutline,
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
        const layout = this._sheets[this._sheetCount - 1]?.layout;
        const totals = this._streamTableTotals.get(this._streamRowNum - 1);
        if (totals) row = withTableTotals(row, totals);
        buffer.writeString(`<row r="${this._streamRowNum}"${xlsxRowAttributes(this._streamRowNum - 1, layout)}>`);
        let prevCol = -1;
        for (let c = 0; c < row.length; c++) {
            const rangeStyle = layout?.styleRanges.styleFor(this._streamRowNum - 1, c);
            prevCol = writeXlsxCell(
                buffer,
                this,
                row[c],
                c,
                prevCol,
                this._streamColLetters[c],
                this._streamRowNum,
                1,
                2,
                rangeStyle,
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
            sheetProtection: sheet.layout?.sheetProtection,
            autoFilterRef: sheet.layout?.autoFilterRef,
            tableParts: sheet.tableParts,
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
        this._streamTableTotals = new Map();
    }

    // Batch API ---------------------------------------------------------

    writeSheet(
        rows: CellValue[][],
        headers: string[] | null = null,
        options: BrowserSheetOptions = {},
    ): void {
        const doAutofilter = options.doAutofilter !== false;
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

        const layout = this._resolveLayout(
            options, headers !== null, columnCount, headers, totalRows,
        );
        sheet.layout = layout;

        const totalsOverrides = tableTotalsOverrides(layout.tables);
        if (totalsOverrides.length > 0) {
            const headerOffset = headers ? 1 : 0;
            const byRow = new Map<number, TableTotalsOverride[]>();
            for (const override of totalsOverrides) {
                const index = override.row - headerOffset;
                if (index < 0 || index >= rows.length) {
                    throw new Error(
                        `Table totals row ${override.row + 1} is outside the written rows `
                        + `(the sheet has ${totalRows} rows)`,
                    );
                }
                const list = byRow.get(index) ?? [];
                list.push(override);
                byRow.set(index, list);
            }
            for (const [index, list] of byRow) {
                rows[index] = withTableTotals(rows[index], list);
            }
        }

        const colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            rows,
            options,
            doAutofilter && headers !== null,
            layout.headerStyle ?? undefined,
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
            freezePanes: layout.freezePanes,
            showGridLines: layout.showGridLines,
            showRowColHeaders: layout.showRowColHeaders,
            showZeroValues: layout.showZeroValues,
            rightToLeft: layout.rightToLeft,
            zoom: layout.zoom,
            tabColor: layout.tabColor,
            hiddenColumns: layout.hiddenColumns,
            columnOutline: layout.columnOutline,
        });

        let rowNum = 0;
        if (headers) {
            rowNum++;
            buffer.writeString(`<row r="${rowNum}"${xlsxRowAttributes(0, layout)}>`);
            const headerStyle = this._headerStyleIndex(layout);
            const baseHeader = layout.headerStyle ?? { font: { bold: true } };
            for (let c = 0; c < headers.length; c++) {
                const rangeStyle = layout.styleRanges.styleFor(0, c);
                const style = rangeStyle
                    ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle))
                    : headerStyle;
                this.writeStringCell(buffer, headers[c], '', style);
            }
            buffer.writeString('</row>');
        }

        for (const row of rows) {
            rowNum++;
            const sheetRow = rowNum - 1;
            buffer.writeString(`<row r="${rowNum}"${xlsxRowAttributes(sheetRow, layout)}>`);
            let prevCol = -1;
            for (let c = 0; c < row.length; c++) {
                const rangeStyle = layout.styleRanges.styleFor(sheetRow, c);
                prevCol = writeXlsxCell(
                    buffer, this, row[c], c, prevCol, colLetters[c], rowNum, 1, 2, rangeStyle,
                );
            }
            buffer.writeString('</row>');
        }

        this._applySheetFeatures(sheet, options, layout);
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
            sheetProtection: layout.sheetProtection,
            autoFilterRef: layout.autoFilterRef,
            tableParts: sheet.tableParts,
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
        const tableParts = this._sheets.flatMap(sheet =>
            (sheet.tableParts ?? []).map(part => ({ path: part.path })),
        );
        this._zip.addFile('[Content_Types].xml', buildXlsxContentTypes(this._sheets, commentParts, tableParts));
        const rels = buildXlsxRels(this._sheets);
        this._zip.addFile('_rels/.rels', rels.global);
        this._zip.addFile('xl/_rels/workbook.xml.rels', rels.workbook);
        return this._zip.toBlob();
    }
}

export default BrowserXlsxWriter;
