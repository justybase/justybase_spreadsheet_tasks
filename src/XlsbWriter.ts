import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import archiver from 'archiver';
import { Readable } from 'stream';
import { BigBuffer } from './BigBuffer';
import {
    CellValue,
    ExcelError,
    CellStyle,
    FormulaCell,
    StyleColor,
    getCellStyle,
    getFormula,
    isRichText,
    mergeCellStyles,
    unwrapCell,
} from './Formats';
import { StreamingSheetState } from './StreamingSheetState';
import {
    sanitizeSheetName,
    initColWidths,
    applyHeaderWidths,
    updateColWidthsFromRows,
    assertColumnWidth,
    defaultColWidth,
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
    timeToOaDate,
    dateHasTime,
    AUTOFIT_SAMPLE_ROWS,
    resolveAutofitRows,
    assertCompressionLevel,
    DEFAULT_COMPRESSION_LEVEL,
} from './writerHelpers';
import {
    writeXlsbBool,
    writeXlsbDouble,
    writeXlsbError,
    writeXlsbInlineString,
    writeXlsbIsst,
    writeXlsbRkInteger,
    writeXlsbRowHeader,
    writeXlsbSharedStrings,
    XlsbSharedString,
    XLSB_STYLES_TEMPLATE,
} from './writers/core/xlsbRecords';
import { XlsbStyleRegistry } from './writers/core/xlsbStyles';
import { buildXlsbFormulaRecord } from './writers/core/xlsbFormula';
import {
    applyXlsbSheetProtection,
    buildXlsbColumnRecords,
    buildXlsbLegacyDrawingRecord,
    buildXlsbPaneRecord,
    buildXlsbSelectionRecord,
    xlsbColorPayload,
} from './writers/core/xlsbLayout';
import {
    buildXlsbHyperlinkRecord,
    buildXlsbMergeRecords,
    buildXlsbValidationRecords,
    parseXlsbRange,
} from './writers/core/xlsbSheetFeatures';
import {
    CommentSpec,
    ConditionalFormatSpec,
    DataValidationSpec,
    HyperlinkSpec,
    SheetLayoutOptions,
    XlsxSheetFeatures,
} from './writers/core/xlsxParts';
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
    validateRowHeight,
    validateZoom,
    tableTotalsOverrides,
    TableTotalsOverride,
} from './writers/core/sheetLayout';
import {
    XlsbDxfRegistry,
    buildXlsbConditionalFormatRecords,
} from './writers/core/xlsbConditionalFormat';
import { buildXlsbCommentsPart, buildXlsbCommentsVml } from './writers/core/xlsbComments';
import { buildXlsbListParts, buildXlsbTablePart } from './writers/core/xlsbTable';
import { escapeXmlText } from './xmlUtils';
import { XLSX_THEME_XML } from './writers/core/theme';

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

interface ResolvedXlsbLayout {
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
    tabColor?: StyleColor;
    sheetProtection?: SheetProtectionSpec;
    autoFilterRef?: string;
    tables: NormalizedTableSpec[];
}

interface SheetInfo {
    name: string;
    pathInArchive: string;
    hidden: boolean;
    nameInArchive: string;
    sheetId: number;
    filterHeaderRange: string | null;
    filterData?: {
        startRow: number;
        endRow: number;
        startColumn: number;
        endColumn: number;
    };
    merges?: string[];
    dataValidations?: DataValidationSpec[];
    conditionalFormats?: ConditionalFormatSpec[];
    cfRecords?: Uint8Array;
    comments?: CommentSpec[];
    commentRId?: string;
    commentVmlRId?: string;
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string; row: number; col: number }>;
    layout?: ResolvedXlsbLayout;
    tableParts?: Array<{ id: number; path: string; data: Uint8Array; rId: string }>;
    extraParts?: Array<{ name: string; data: Uint8Array }>;
}

/** Options accepted by the XLSB streaming and batch sheet APIs. */
export interface XlsbSheetOptions extends XlsxSheetFeatures, SheetLayoutOptions {
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
 * High-performance XLSB (Excel Binary Workbook) writer.
 * Prefer this format for large datasets — typically faster and smaller than XLSX.
 */
export class XlsbWriter {
    private output: fs.WriteStream;
    private archive: ReturnType<typeof archiver>;
    private readonly _targetPath: string;
    private readonly _tempPath: string;
    private _outputError: Error | null = null;
    private _finalized: boolean = false;

    private sheetCount: number = 0;
    private sheetList: SheetInfo[] = [];
    private sstDic: Map<string, number> = new Map();
    private sstRichDic: Map<string, number> = new Map();
    private sstValues: XlsbSharedString[] = [];
    private sstCntUnique: number = 0;
    private sstCntAll: number = 0;
    private _autofilterIsOn: boolean = false;
    private _nextTableId: number = 1;
    private readonly _dxfs = new XlsbDxfRegistry(
        format => this._styles.registerNumberFormat(format),
    );

    private _oaEpoch: number;
    private readonly _useSharedStrings: boolean;
    private readonly _styles = new XlsbStyleRegistry();

    /** Worksheet binary template (BrtBeginSheet … dimension / views / cols placeholders). */
    private _sheet1Bytes: Buffer;
    private _workbookBinStart: Buffer;
    private _workbookBinMiddle: Buffer;
    private _workbookBinEnd: Buffer;
    /** Default styles.bin template; extended when custom number formats are registered. */
    private _stylesBin: Buffer;
    private _binaryIndexBin: Buffer;
    private _rRkIntegerLowerLimit: number;
    private _rRkIntegerUpperLimit: number;
    private _autoFilterStartBytes: Buffer;
    private _autoFilterEndBytes: Buffer;
    private _stickHeaderA1bytes: Buffer;
    /** Excel 2016+ autofilter compatibility stubs. */
    private _magicFilterExcel2016Fix0: Buffer;
    private _magicFilterExcel2016Fix1: Buffer;
    private _magicFilterExcel2016Fix2: Buffer;

    private readonly stream = new StreamingSheetState();
    private _streamPending: BigBuffer | null = null;
    private _streamHeadFlushed: boolean = true;
    private _streamHasHeaders: boolean = false;
    private _streamDataRows: number = 0;
    private _streamAutofitRows: number = AUTOFIT_SAMPLE_ROWS;
    private _streamTableTotals: Map<number, TableTotalsOverride[]> = new Map();


    /**
     * @param filePath Destination `.xlsb` path on disk.
     */
    /**
     * @param filePath Destination `.xlsb` path on disk.
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

        this._sheet1Bytes = Buffer.from([
            0x81, 0x01, 0x00, 0x93, 0x01, 0x17, 0xCB, 0x04,
            0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0x00, 0x00, 0x00, 0x00, 0x94, 0x01, 0x10,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x85, 0x01, 0x00, 0x89, 0x01, 0x1E, 0xDC, 0x03,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00,
            0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x98, 0x01, 0x24, 0x03,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x8A, 0x01, 0x00, 0x86, 0x01,
            0x00, 0x25, 0x06, 0x01, 0x00, 0x02, 0x0E, 0x00,
            0x80, 0x95, 0x08, 0x02, 0x05, 0x00, 0x26, 0x00,
            0xE5, 0x03, 0x0C, 0xFF, 0xFF, 0xFF, 0xFF, 0x08,
            0x00, 0x2C, 0x01, 0x00, 0x00, 0x00, 0x00, 0x91,
            0x01, 0x00, 0x25, 0x06, 0x01, 0x00, 0x02, 0x0E,
            0x00, 0x80, 0x80, 0x08, 0x02, 0x05, 0x00, 0x26,
            0x00, 0x00, 0x19, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x2C, 0x01, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x07, 0x0C, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x92, 0x01, 0x00, 0x97, 0x04, 0x42,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00,
            0xDD, 0x03, 0x02, 0x10, 0x00, 0xDC, 0x03, 0x30,
            0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0xE6, 0x3F,
            0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0xE6, 0x3F,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE8, 0x3F,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE8, 0x3F,
            0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xD3, 0x3F,
            0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xD3, 0x3F,
            0x25, 0x06, 0x01, 0x00, 0x00, 0x10, 0x00, 0x80,
            0x80, 0x18, 0x10, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x26, 0x00, 0x82, 0x01, 0x00
        ]);

        this._workbookBinStart = Buffer.from([
            0x83, 0x01, 0x00, 0x80, 0x01, 0x32, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x78, 0x00, 0x6C, 0x00, 0x01, 0x00, 0x00, 0x00,
            0x37, 0x00, 0x01, 0x00, 0x00, 0x00, 0x36, 0x00, 0x05, 0x00, 0x00, 0x00, 0x32, 0x00, 0x34, 0x00, 0x33, 0x00, 0x32, 0x00, 0x36, 0x00, 0x99, 0x01, 0x0C, 0x20, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x25, 0x06, 0x01, 0x00, 0x03, 0x0F, 0x00, 0x80, 0x97, 0x10, 0x34, 0x18, 0x00, 0x00, 0x00, 0x43, 0x00, 0x3A, 0x00, 0x5C, 0x00, 0x73, 0x00, 0x71, 0x00, 0x6C, 0x00, 0x73, 0x00, 0x5C, 0x00,
            0x54, 0x00, 0x65, 0x00, 0x73, 0x00, 0x74, 0x00, 0x79, 0x00, 0x5A, 0x00, 0x61, 0x00, 0x70, 0x00, 0x69, 0x00, 0x73, 0x00, 0x75, 0x00, 0x58, 0x00, 0x6C, 0x00, 0x73, 0x00, 0x62, 0x00, 0x5C, 0x00, 0x26, 0x00,
            0x25, 0x06, 0x01, 0x00, 0x00, 0x10, 0x00, 0x80, 0x81, 0x18, 0x82, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x2F, 0x00, 0x00, 0x00, 0x31, 0x00, 0x33, 0x00, 0x5F, 0x00, 0x6E, 0x00, 0x63, 0x00,
            0x72, 0x00, 0x3A, 0x00, 0x31, 0x00, 0x5F, 0x00, 0x7B, 0x00, 0x31, 0x00, 0x36, 0x00, 0x35, 0x00, 0x30, 0x00, 0x38, 0x00, 0x44, 0x00, 0x36, 0x00, 0x39, 0x00, 0x2D, 0x00, 0x43, 0x00, 0x46, 0x00, 0x38, 0x00,
            0x37, 0x00, 0x2D, 0x00, 0x34, 0x00, 0x37, 0x00, 0x36, 0x00, 0x39, 0x00, 0x2D, 0x00, 0x38, 0x00, 0x34, 0x00, 0x35, 0x00, 0x36, 0x00, 0x2D, 0x00, 0x44, 0x00, 0x34, 0x00, 0x41, 0x00, 0x34, 0x00, 0x30, 0x00,
            0x31, 0x00, 0x31, 0x00, 0x33, 0x00, 0x31, 0x00, 0x35, 0x00, 0x36, 0x00, 0x37, 0x00, 0x7D, 0x00, 0x2F, 0x00, 0x00, 0x00, 0x2F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x87, 0x01, 0x00, 0x25, 0x06, 0x01, 0x00, 0x02, 0x10, 0x00, 0x80, 0x80, 0x18, 0x10, 0x00, 0x00, 0x00, 0x00, 0x0D, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x9E, 0x01, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x9E, 0x16, 0x00, 0x00, 0xB4, 0x69, 0x00, 0x00, 0xE8, 0x26, 0x00, 0x00, 0x58, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x78, 0x88, 0x01, 0x00,
            0x8F, 0x01, 0x00
        ]);

        this._workbookBinMiddle = Buffer.from([0x90, 0x01, 0x00]);

        this._workbookBinEnd = Buffer.from([
            0x9D, 0x01, 0x1A, 0x35, 0xEA, 0x02, 0x00, 0x01, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0xFC, 0xA9, 0xF1, 0xD2, 0x4D, 0x62, 0x50, 0x3F, 0x01,
            0x00, 0x00, 0x00, 0x6A, 0x00, 0x9B, 0x01, 0x01, 0x00, 0x23, 0x04, 0x03, 0x0F, 0x00, 0x00, 0xAB, 0x10, 0x01, 0x01, 0x24, 0x00, 0x84, 0x01, 0x00
        ]);

        this._stylesBin = Buffer.from(XLSB_STYLES_TEMPLATE);

        this._binaryIndexBin = Buffer.from([
            0x2A, 0x18, 0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x95,
            0x02, 0x00
        ]);

        this._rRkIntegerLowerLimit = -1 << 29;
        this._rRkIntegerUpperLimit = (1 << 29) - 1;

        this._autoFilterStartBytes = Buffer.from([0xA1, 0x01, 0x10]);
        this._autoFilterEndBytes = Buffer.from([0xA2, 0x01, 0x00]);

        this._stickHeaderA1bytes = Buffer.from([
            0x97, 0x01, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0xF0, 0x3F, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x03
        ]);

        this._magicFilterExcel2016Fix0 = Buffer.from([0xE1, 0x02, 0x00, 0xE5, 0x02, 0x00, 0xEA, 0x02]);
        this._magicFilterExcel2016Fix1 = Buffer.from([
            0x27, 0x46, 0x21, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x0F, 0x00, 0x00, 0x00, 0x5F,
            0x00, 0x46, 0x00, 0x69, 0x00, 0x6C, 0x00, 0x74, 0x00, 0x65, 0x00, 0x72, 0x00, 0x44, 0x00, 0x61,
            0x00, 0x74, 0x00, 0x61, 0x00, 0x62, 0x00, 0x61, 0x00, 0x73, 0x00, 0x65, 0x00, 0x0F, 0x00, 0x00,
            0x00, 0x3B, 0xFF, 0x00
        ]);
        this._magicFilterExcel2016Fix2 = Buffer.from([0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF]);
    }

    /** @deprecated Internal; kept for unit-test access. Prefer sanitizeSheetName(). */
    private _sanitizeSheetName(name: string): string {
        return sanitizeSheetName(name, this.sheetCount);
    }

    private _buildStylesBin(): Uint8Array {
        return this._styles.toStylesBin(this._stylesBin);
    }

    /**
     * Resolve column widths: explicit `columnWidths` win, otherwise plain
     * autofit from headers and up to 200 sample rows, capped at 200.
     */
    private _resolveColumnWidths(
        columnCount: number,
        headers: string[] | null | undefined,
        rows: unknown[][] | undefined,
        options: XlsbSheetOptions,
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

    private _resolveLayout(
        options: XlsbSheetOptions,
        hasHeaders: boolean,
        columnCount: number,
        headers: string[] | null | undefined,
        rowCount: number = 0,
    ): ResolvedXlsbLayout {
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
            headerStyle: options.headerStyle ?? null,
            styleRanges: new StyleRangeSet(options.styles),
            rowHeights,
            hiddenRows,
            hiddenColumns,
            rowOutline: buildOutline(rowGroups, rowMax),
            columnOutline: buildOutline(columnGroups, columnMax),
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

    private _headerStyleIndex(layout: ResolvedXlsbLayout): number {
        if (!layout.headerStyle) return 3;
        return this._styles.register(mergeCellStyles({ font: { bold: true } }, layout.headerStyle));
    }

    /** Store sheet features and build the worksheet extra parts. */
    private _applySheetFeatures(sheet: SheetInfo, features: XlsbSheetOptions): void {
        if (features.merges) sheet.merges = features.merges;
        if (features.dataValidations) sheet.dataValidations = features.dataValidations;

        if (features.conditionalFormats && features.conditionalFormats.length > 0) {
            sheet.conditionalFormats = features.conditionalFormats;
            // Priorities are scoped to one sheet, so every sheet starts at 1.
            sheet.cfRecords = buildXlsbConditionalFormatRecords(
                features.conditionalFormats,
                this._dxfs,
                1,
            );
        }

        let nextRelId = 2; // rId1 is the worksheet binary index
        if (features.hyperlinks && features.hyperlinks.length > 0) {
            sheet.hyperlinks = features.hyperlinks.map(spec => {
                const range = parseXlsbRange(spec.ref);
                const internal = spec.target === undefined;
                const rId = internal ? undefined : `rId${nextRelId++}`;
                return { spec, rId, row: range.r1, col: range.c1 };
            });
        }

        const layout = sheet.layout;
        if (layout && layout.tables.length > 0) {
            sheet.tableParts = [];
            for (let index = 0; index < layout.tables.length; index++) {
                const table = layout.tables[index];
                const id = this._nextTableId - layout.tables.length + index;
                const path = `xl/tables/table${id}.bin`;
                const data = buildXlsbTablePart(table, id);
                sheet.tableParts.push({ id, path, data, rId: `rId${nextRelId++}` });
                sheet.extraParts = [...(sheet.extraParts ?? []), { name: path, data }];
            }
        }

        if (features.comments && features.comments.length > 0) {
            sheet.comments = features.comments;
            sheet.commentRId = `rId${nextRelId++}`;
            sheet.commentVmlRId = `rId${nextRelId++}`;
            const commentPath = `xl/comments/comment${sheet.sheetId}.bin`;
            const vmlPath = `xl/drawings/commentsDrawing${sheet.sheetId}.vml`;
            sheet.extraParts = [
                ...(sheet.extraParts ?? []),
                { name: commentPath, data: buildXlsbCommentsPart(features.comments, commentPath) },
                {
                    name: vmlPath,
                    data: new TextEncoder().encode(buildXlsbCommentsVml(features.comments)),
                },
            ];
        }
    }

    /** Merged cells, validations, hyperlinks and conditional formats. */
    private _writeFeatureRecords(bigBuf: BigBuffer, sheet: SheetInfo): void {
        if (sheet.merges && sheet.merges.length > 0) {
            bigBuf.write(buildXlsbMergeRecords(sheet.merges));
        }
        if (sheet.dataValidations && sheet.dataValidations.length > 0) {
            bigBuf.write(buildXlsbValidationRecords(sheet.dataValidations));
        }
        for (const link of sheet.hyperlinks ?? []) {
            bigBuf.write(buildXlsbHyperlinkRecord(link.row, link.col, link.spec, link.rId));
        }
        if (sheet.cfRecords && sheet.cfRecords.length > 0) {
            bigBuf.write(sheet.cfRecords);
        }
    }

    /** Write the template's tail plus table/list/comment records. */
    private _writeSheetFooterRecords(
        bigBuf: BigBuffer,
        sheet: SheetInfo,
        sheetHeader: Buffer,
    ): void {
        bigBuf.write(sheetHeader.subarray(290, 346));
        const relIds = (sheet.tableParts ?? []).map(part => part.rId);
        if (relIds.length > 0) bigBuf.write(buildXlsbListParts(relIds));
        if (sheet.commentVmlRId !== undefined) {
            bigBuf.write(buildXlsbLegacyDrawingRecord(sheet.commentVmlRId));
        }
        bigBuf.write(sheetHeader.subarray(346));
    }

    /** Write one cell using the current style registry and formula compiler. */
    private _writeCell(bigBuf: BigBuffer, raw: CellValue, col: number, extraStyle?: CellStyle | null): void {
        if (raw === null || raw === undefined) return;

        const cellStyle = getCellStyle(raw);
        const style = extraStyle ? mergeCellStyles(cellStyle, extraStyle) : cellStyle;
        let styleNum = style !== null ? this._styles.register(style) : 0;

        const formula = getFormula(raw);
        if (formula !== null) {
            const rawCached = (raw as FormulaCell).value;
            let cached: string | number | boolean | ExcelError | null | undefined;
            if (rawCached instanceof Date) {
                cached = timeToOaDate(rawCached.getTime(), this._oaEpoch);
                if (styleNum === 0) styleNum = dateHasTime(rawCached) ? 1 : 2;
            } else if (isRichText(rawCached)) {
                cached = rawCached.toString();
            } else if (typeof rawCached === 'bigint') {
                cached = rawCached.toString();
            } else {
                cached = rawCached;
            }
            bigBuf.write(buildXlsbFormulaRecord(col, styleNum, formula, cached));
            return;
        }

        const val = unwrapCell(raw);
        if (val === null || val === undefined) return;
        if (isRichText(val)) {
            this._writeRichText(bigBuf, val, col, styleNum);
            return;
        }

        if (typeof val === 'number') {
            if (Number.isInteger(val)) {
                if (val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
                    this.writeRkNumberInteger(bigBuf, val, col, styleNum);
                } else {
                    this.writeDouble(bigBuf, val, col, styleNum);
                }
            } else {
                this.writeDouble(bigBuf, val, col, styleNum);
            }
        } else if (val instanceof ExcelError) {
            this.writeError(bigBuf, val.code, col, styleNum);
        } else if (typeof val === 'bigint') {
            this.writeString(bigBuf, val.toString(), col, false, styleNum);
        } else if (typeof val === 'boolean') {
            this.writeBool(bigBuf, val, col, styleNum);
        } else if (val instanceof Date) {
            const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
            // Base XF 1 carries the date-time format, XF 2 the date-only one.
            const dateStyle = dateHasTime(val) ? 1 : 2;
            this.writeDouble(bigBuf, oaDate, col, styleNum !== 0 ? styleNum : dateStyle);
        } else {
            this.writeString(bigBuf, val.toString(), col, false, styleNum);
        }
    }

/**
     * Register a worksheet (metadata only). Prefer `writeSheet` / `startSheet` for data.
     * @param sheetName Display name (sanitized to Excel rules)
     * @param hidden Whether the sheet is hidden in Excel
     */
    addSheet(sheetName: string, hidden: boolean = false): void {
        const sanitizedName = sanitizeSheetName(sheetName, this.sheetCount);
        this.sheetCount++;
        this.sheetList.push({
            name: sanitizedName,
            pathInArchive: `xl/worksheets/sheet${this.sheetCount}.bin`,
            hidden: hidden,
            nameInArchive: `sheet${this.sheetCount}.bin`,
            sheetId: this.sheetCount,
            filterHeaderRange: null
        });
    }

    /**
     * Start a new sheet in streaming mode. Call writeRow() for each row, then endSheet() when done.
     * @param sheetName Name of the sheet
     * @param columnCount Number of columns in the sheet
     * @param headers Optional header row
     * @param options Optional settings (hidden, doAutofilter)
     */
    startSheet(
        sheetName: string,
        columnCount: number,
        headers?: string[],
        options: XlsbSheetOptions = {}
    ): void {
        assertColumnCount(columnCount);
        if (headers !== undefined) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        const { hidden = false, doAutofilter = true } = options;

        // Add sheet metadata
        this.addSheet(sheetName, hidden);
        const sheet = this.sheetList[this.sheetCount - 1];
        const layout = this._resolveLayout(options, headers !== undefined, columnCount, headers);
        sheet.layout = layout;
        this._streamTableTotals = new Map();
        for (const override of tableTotalsOverrides(layout.tables)) {
            const list = this._streamTableTotals.get(override.row) ?? [];
            list.push(override);
            this._streamTableTotals.set(override.row, list);
        }
        this._applySheetFeatures(sheet, options);

        this.stream.begin(columnCount, doAutofilter && headers !== undefined, new BigBuffer());

        this.stream.colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            undefined,
            options,
            doAutofilter && headers !== undefined,
            layout.headerStyle ?? undefined,
        );

        // The sheet head carries the widths, so it is delayed until the first
        // AUTOFIT_SAMPLE_ROWS rows have been measured (200 by default).
        this._streamPending = new BigBuffer();
        this._streamHeadFlushed = false;
        this._streamHasHeaders = headers !== undefined;
        this._streamDataRows = 0;
        this._streamAutofitRows = resolveAutofitRows(options.autofitRows);

        // Write header row if provided
        if (headers) {
            const pending = this._streamPending;
            this.createRowHeader(
                pending, this.stream.rowNum, this.stream.startCol, this.stream.endCol,
                this._rowHeaderOptions(layout, 0),
            );
            const headerStyle = this._headerStyleIndex(layout);
            const baseHeader = mergeCellStyles({ font: { bold: true } }, layout.headerStyle);
            for (let c = 0; c < headers.length; c++) {
                const rangeStyle = layout.styleRanges.styleFor(0, c);
                const style = rangeStyle
                    ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle))
                    : headerStyle;
                this.writeString(pending, headers[c], c, false, style);
            }
            this.stream.rowNum++;
        }
        if (this._streamAutofitRows === 0) this._flushStreamHead();
    }

    /** Write the delayed sheet head and everything buffered before it. */
    private _flushStreamHead(): void {
        if (this._streamHeadFlushed) return;
        const bigBuf = this.stream.assertStreaming();
        const pending = this._streamPending;

        const layout = this.sheetList[this.sheetCount - 1]?.layout;
        const sheetHeader = Buffer.from(this._sheet1Bytes);
        applyXlsbSheetProtection(sheetHeader, layout?.sheetProtection);
        sheetHeader.writeInt32LE(this.stream.startCol, 40);
        sheetHeader.writeInt32LE(this.stream.endCol - 1, 44);

        // Tab colour lives in the BrtWsProp payload (bytes 9..16).
        if (layout?.tabColor !== undefined) {
            Buffer.from(xlsbColorPayload(layout.tabColor)).copy(sheetHeader, 9);
        }

        // BrtBeginWsView flags: fDspGrid 0x04, fDspRwCol 0x08,
        // fDspZeros 0x10, fRightToLeft 0x20, fSelected 0x40.
        let viewFlags = this.sheetCount === 1 ? 0xDC : 0x9C;
        if (layout?.showGridLines === false) viewFlags &= ~0x04;
        if (layout?.showRowColHeaders === false) viewFlags &= ~0x08;
        if (layout?.showZeroValues === false) viewFlags &= ~0x10;
        if (layout?.rightToLeft) viewFlags |= 0x20;
        sheetHeader[54] = viewFlags;

        bigBuf.write(sheetHeader.subarray(0, 55));

        let viewTail = Buffer.from(sheetHeader.subarray(55, 84));
        if (layout?.zoom !== undefined) {
            const scale = Buffer.alloc(4);
            scale.writeUInt16LE(layout.zoom, 0);
            scale.writeUInt16LE(layout.zoom, 2);
            viewTail = Buffer.concat([viewTail.subarray(0, 17), scale, viewTail.subarray(21)]);
        }
        bigBuf.write(viewTail);

        bigBuf.write(buildXlsbPaneRecord(layout?.freezePanes ?? null));
        bigBuf.write(buildXlsbSelectionRecord(layout?.freezePanes ?? null));
        bigBuf.write(sheetHeader.subarray(123, 157));
        bigBuf.write(Buffer.from([
            layout?.rowOutline.maxLevel ?? 0,
            layout?.columnOutline.maxLevel ?? 0,
        ]));

        const widths = this.stream.colWidths.map(width => defaultColWidth(width));
        bigBuf.write(buildXlsbColumnRecords({
            startCol: this.stream.startCol,
            endCol: this.stream.endCol,
            widths,
            hiddenColumns: layout?.hiddenColumns ?? [],
            outline: layout?.columnOutline,
        }));

        bigBuf.write(sheetHeader.subarray(159, 175));
        bigBuf.write(Buffer.from([38, 0]));

        if (pending) {
            for (const chunk of pending.getChunks()) bigBuf.write(chunk);
        }
        this._streamPending = null;
        this._streamHeadFlushed = true;
    }

    private _rowHeaderOptions(layout: ResolvedXlsbLayout | undefined, rowIndex: number): import('./writers/core/xlsbRecords').XlsbRowHeaderOptions {
        if (!layout) return {};
        const options: import('./writers/core/xlsbRecords').XlsbRowHeaderOptions = {};
        const height = layout.rowHeights[rowIndex];
        if (height !== undefined) options.heightTwips = Math.round(height * 20);
        if (layout.hiddenRows.has(rowIndex)) options.hidden = true;
        if (rowIndex < layout.rowOutline.level.length) {
            options.outlineLevel = layout.rowOutline.level[rowIndex];
            if (layout.rowOutline.hidden[rowIndex]) options.hidden = true;
            if (layout.rowOutline.collapsed[rowIndex]) options.collapsed = true;
        }
        return options;
    }

    private _writeStreamRow(bigBuf: BigBuffer, row: CellValue[]): void {
        assertRowCount(this.stream.rowNum + 1);
        const layout = this.sheetList[this.sheetCount - 1]?.layout;
        const totals = this._streamTableTotals.get(this.stream.rowNum);
        if (totals) row = withTableTotals(row, totals);
        this.createRowHeader(
            bigBuf,
            this.stream.rowNum,
            this.stream.startCol,
            this.stream.endCol,
            this._rowHeaderOptions(layout, this.stream.rowNum),
        );
        for (let c = 0; c < row.length; c++) {
            this._writeCell(bigBuf, row[c], c, layout?.styleRanges.styleFor(this.stream.rowNum, c));
        }
        this.stream.rowNum++;
    }

    /**
     * Write a single row in streaming mode. Must be called between startSheet() and endSheet().
     * @param row Array of cell values
     */
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

    /**
     * Finalize the current streaming sheet and add it to the archive.
     * Must be called after startSheet() and writeRow() calls.
     */
    endSheet(): void {
        this._flushStreamHead();
        const bigBuf = this.stream.assertStreaming();
        const sheet = this.sheetList[this.sheetCount - 1];
        const sheetHeader = Buffer.from(this._sheet1Bytes);
        applyXlsbSheetProtection(sheetHeader, sheet.layout?.sheetProtection);

        // Write sheet footer sections
        bigBuf.write(sheetHeader.subarray(218, 290));

        // Write autofilter if enabled
        if (this.stream.doAutofilter) {
            this._autofilterIsOn = true;
            const endRow = this.stream.rowNum;

            bigBuf.write(this._autoFilterStartBytes);

            const rowBuf = Buffer.alloc(8);
            rowBuf.writeInt32LE(0, 0);
            rowBuf.writeInt32LE(endRow - 1, 4);
            bigBuf.write(rowBuf);

            const colBuf = Buffer.alloc(8);
            colBuf.writeInt32LE(this.stream.startCol, 0);
            colBuf.writeInt32LE(this.stream.endCol - 1, 4);
            bigBuf.write(colBuf);

            bigBuf.write(this._autoFilterEndBytes);

            // Store filter data for later
            sheet.filterData = {
                startRow: 0,
                endRow: endRow - 1,
                startColumn: this.stream.startCol,
                endColumn: this.stream.endCol - 1
            };
        }

        this._writeFeatureRecords(bigBuf, sheet);
        this._writeSheetFooterRecords(bigBuf, sheet, sheetHeader);

        // BrtWsDim was written with row 0; patch the inclusive last row now
        // that the row count is known (fixed offset 36 in the sheet header).
        bigBuf.patchInt32LE(36, Math.max(0, this.stream.rowNum - 1));

        // Add to archive
        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: sheet.pathInArchive
        });

        // Reset streaming state
        this.stream.end();
        this._streamPending = null;
        this._streamHeadFlushed = true;
        this._streamHasHeaders = false;
        this._streamDataRows = 0;
        this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
        this._streamTableTotals = new Map();
    }


    /**
     * Write an entire sheet in one call (batch mode).
     *
     * The third argument keeps the legacy `doAutofilter` boolean; pass
     * {@link XlsbSheetOptions} instead to add merges, hyperlinks, validations
     * and width control.
     */
    writeSheet(
        rows: CellValue[][],
        headers: string[] | null = null,
        doAutofilterOrOptions: boolean | XlsbSheetOptions = true,
        extraOptions: XlsbSheetOptions = {},
    ): void {
        let doAutofilter = true;
        let options: XlsbSheetOptions;
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
        const totalRowsCheck = rows.length + (headers ? 1 : 0);
        assertRowCount(totalRowsCheck);
        if (headers) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        for (let r = 0; r < rows.length; r++) {
            assertRowFitsColumns('row', rows[r].length, columnCount);
        }

        const sheet = this.sheetList[this.sheetCount - 1];
        if (!sheet) {
            throw new Error('Call addSheet() before writeSheet()');
        }
        const layout = this._resolveLayout(
            options,
            headers !== null && headers !== undefined,
            columnCount,
            headers,
            totalRowsCheck,
        );
        sheet.layout = layout;
        this._applySheetFeatures(sheet, options);

        const totalsOverrides = tableTotalsOverrides(layout.tables);
        if (totalsOverrides.length > 0) {
            const headerOffset = headers ? 1 : 0;
            const byRow = new Map<number, TableTotalsOverride[]>();
            for (const override of totalsOverrides) {
                const index = override.row - headerOffset;
                if (index < 0 || index >= rows.length) {
                    throw new Error(
                        `Table totals row ${override.row + 1} is outside the written rows `
                        + `(the sheet has ${totalRowsCheck} rows)`,
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
            doAutofilter && headers !== null && headers !== undefined,
            layout.headerStyle ?? undefined,
        );

        const sheetHeader = Buffer.from(this._sheet1Bytes);
        applyXlsbSheetProtection(sheetHeader, layout.sheetProtection);
        const startCol = 0;
        const endCol = columnCount;

        sheetHeader.writeInt32LE(startCol, 40);
        sheetHeader.writeInt32LE(endCol - 1, 44);

        if (layout.tabColor !== undefined) {
            Buffer.from(xlsbColorPayload(layout.tabColor)).copy(sheetHeader, 9);
        }

        let viewFlags = this.sheetCount === 1 ? 0xDC : 0x9C;
        if (!layout.showGridLines) viewFlags &= ~0x04;
        if (!layout.showRowColHeaders) viewFlags &= ~0x08;
        if (!layout.showZeroValues) viewFlags &= ~0x10;
        if (layout.rightToLeft) viewFlags |= 0x20;
        sheetHeader[54] = viewFlags;

        bigBuf.write(sheetHeader.subarray(0, 55));
        let viewTail = Buffer.from(sheetHeader.subarray(55, 84));
        if (layout.zoom !== undefined) {
            const scale = Buffer.alloc(4);
            scale.writeUInt16LE(layout.zoom, 0);
            scale.writeUInt16LE(layout.zoom, 2);
            viewTail = Buffer.concat([viewTail.subarray(0, 17), scale, viewTail.subarray(21)]);
        }
        bigBuf.write(viewTail);
        bigBuf.write(buildXlsbPaneRecord(layout.freezePanes));
        bigBuf.write(buildXlsbSelectionRecord(layout.freezePanes));
        bigBuf.write(sheetHeader.subarray(123, 157));
        bigBuf.write(Buffer.from([
            layout.rowOutline.maxLevel,
            layout.columnOutline.maxLevel,
        ]));

        const widths = colWidths.map(width => defaultColWidth(width));
        bigBuf.write(buildXlsbColumnRecords({
            startCol,
            endCol,
            widths,
            hiddenColumns: layout.hiddenColumns,
            outline: layout.columnOutline,
        }));

        bigBuf.write(sheetHeader.subarray(159, 175));
        bigBuf.write(Buffer.from([38, 0]));

        let rowNum = 0;

        if (headers) {
            this.createRowHeader(
                bigBuf, rowNum, startCol, endCol, this._rowHeaderOptions(layout, 0),
            );
            const headerStyle = this._headerStyleIndex(layout);
            const baseHeader = mergeCellStyles({ font: { bold: true } }, layout.headerStyle);
            for (let c = 0; c < headers.length; c++) {
                const rangeStyle = layout.styleRanges.styleFor(0, c);
                const style = rangeStyle
                    ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle))
                    : headerStyle;
                this.writeString(bigBuf, headers[c], c, false, style);
            }
            rowNum++;
        }

        for (let r = 0; r < rows.length; r++) {
            this.createRowHeader(
                bigBuf, rowNum, startCol, endCol, this._rowHeaderOptions(layout, rowNum),
            );
            const row = rows[r];
            for (let c = 0; c < row.length; c++) {
                this._writeCell(bigBuf, row[c], c, layout.styleRanges.styleFor(rowNum, c));
            }
            rowNum++;
        }

        bigBuf.write(sheetHeader.subarray(218, 290));

        if (doAutofilter && headers) {
            this._autofilterIsOn = true;
            const endRow = rows.length + 1;

            bigBuf.write(this._autoFilterStartBytes);

            const rowBuf = Buffer.alloc(8);
            rowBuf.writeInt32LE(0, 0);
            rowBuf.writeInt32LE(endRow - 1, 4);
            bigBuf.write(rowBuf);

            const colBuf = Buffer.alloc(8);
            colBuf.writeInt32LE(startCol, 0);
            colBuf.writeInt32LE(endCol - 1, 4);
            bigBuf.write(colBuf);

            bigBuf.write(this._autoFilterEndBytes);

            sheet.filterData = {
                startRow: 0,
                endRow: rows.length,
                startColumn: startCol,
                endColumn: endCol - 1
            };
        }

        this._writeFeatureRecords(bigBuf, sheet);
        this._writeSheetFooterRecords(bigBuf, sheet, sheetHeader);

        bigBuf.patchInt32LE(36, Math.max(0, rowNum - 1));

        this.archive.append(Readable.from(bigBuf.getChunks()), { name: sheet.pathInArchive });
    }



    createRowHeader(
        bigBuf: BigBuffer,
        rowNumber: number,
        startCol: number,
        endCol: number,
        options: import('./writers/core/xlsbRecords').XlsbRowHeaderOptions = {},
    ): void {
        writeXlsbRowHeader(bigBuf, rowNumber, startCol, endCol, options);
    }

    writeRkNumberInteger(bigBuf: BigBuffer, val: number, colNum: number, styleNum: number = 0): void {
        writeXlsbRkInteger(bigBuf, val, colNum, styleNum);
    }

    writeDouble(bigBuf: BigBuffer, val: number, colNum: number, styleNum: number = 0): void {
        writeXlsbDouble(bigBuf, val, colNum, styleNum);
    }

    writeBool(bigBuf: BigBuffer, val: boolean, colNum: number, styleNum: number = 0): void {
        writeXlsbBool(bigBuf, val, colNum, styleNum);
    }

    /** BrtCellError: column, style (24-bit) and the BIFF12 error code. */
    writeError(bigBuf: BigBuffer, code: string, colNum: number, styleNum: number = 0): void {
        writeXlsbError(bigBuf, code, colNum, styleNum);
    }

    writeDateTime(bigBuf: BigBuffer, date: Date, colNum: number): void {
        // Use UTC time to avoid timezone conversion issues; handles the
        // fictitious 1900-02-29.
        const oaDate = timeToOaDate(date.getTime(), this._oaEpoch);
        this.writeDouble(bigBuf, oaDate, colNum, dateHasTime(date) ? 1 : 2);
    }

    private _writeRichText(
        bigBuf: BigBuffer,
        value: import('./Formats').RichText,
        colNum: number,
        styleNum: number,
    ): void {
        if (!this._useSharedStrings) {
            // BrtCellSt can only hold plain text; keep the text, drop the runs.
            this.writeString(bigBuf, value.toString(), colNum, false, styleNum);
            return;
        }
        const runs: Array<{ ich: number; ifnt: number }> = [];
        let offset = 0;
        for (const run of value.runs) {
            runs.push({ ich: offset, ifnt: this._styles.registerRichFont(run.style?.font) });
            offset += run.text.length;
        }
        const text = value.toString();
        assertCellTextLength(text);
        const key = JSON.stringify({ text, runs });
        let index = this.sstRichDic.get(key);
        if (index === undefined) {
            index = this.sstValues.length;
            this.sstRichDic.set(key, index);
            this.sstValues.push({ text, runs });
            this.sstCntUnique++;
        }
        this.sstCntAll++;
        writeXlsbIsst(bigBuf, index, colNum, styleNum);
    }

    writeString(bigBuf: BigBuffer, val: string, colNum: number, bolded: boolean = false, styleOverride?: number): void {
        assertCellTextLength(val);
        const styleNum = styleOverride !== undefined ? styleOverride : (bolded ? 3 : 0);

        if (!this._useSharedStrings) {
            writeXlsbInlineString(bigBuf, val, colNum, styleNum);
            return;
        }

        let index = this.sstDic.get(val);
        if (index === undefined) {
            index = this.sstValues.length;
            this.sstDic.set(val, index);
            this.sstValues.push(val);
            this.sstCntUnique++;
        }
        this.sstCntAll++;

        writeXlsbIsst(bigBuf, index, colNum, styleNum);
    }

    private saveSst(): void {
        const bigBuf = new BigBuffer();
        writeXlsbSharedStrings(bigBuf, this.sstValues, this.sstCntAll, this.sstCntUnique);
        this.archive.append(Readable.from(bigBuf.getChunks()), { name: 'xl/sharedStrings.bin' });
    }

    private _writeFilterDefinedName(wbBuffers: Buffer[], sheet: SheetInfo, sheetNum: number): void {
        const filterData = sheet.filterData!;
        const sheetIndex = sheet.sheetId - 1;

        const fix1 = Buffer.alloc(this._magicFilterExcel2016Fix1.length);
        this._magicFilterExcel2016Fix1.copy(fix1);

        const lastIdx = this._magicFilterExcel2016Fix1.length - 2;

        fix1[7] = sheetIndex;
        fix1[lastIdx] = sheetNum;

        wbBuffers.push(fix1);

        const rowBuf = Buffer.alloc(8);
        rowBuf.writeInt32LE(filterData.startRow, 0);
        rowBuf.writeInt32LE(filterData.endRow, 4);
        wbBuffers.push(rowBuf);

        const colBuf = Buffer.alloc(4);
        colBuf.writeInt16LE(filterData.startColumn, 0);
        colBuf.writeInt16LE(filterData.endColumn, 2);
        wbBuffers.push(colBuf);

        wbBuffers.push(this._magicFilterExcel2016Fix2);
    }

    /**
     * Finalize the ZIP package and close the output stream.
     * Must be awaited before the file is fully written.
     */
    finalize(): Promise<void> {
        return new Promise((resolve, reject) => {
            if (this._finalized) {
                reject(new Error('XlsbWriter: already finalized'));
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
                this.saveSst();

                this.archive.append(
                    Buffer.from(this._styles.toStylesBin(this._stylesBin, this._dxfs.toRecords())),
                    { name: 'xl/styles.bin' },
                );
                for (const sheet of this.sheetList) {
                    for (const part of sheet.extraParts ?? []) {
                        this.archive.append(Buffer.from(part.data), { name: part.name });
                    }
                }

                const wbBuffers: Buffer[] = [];
                wbBuffers.push(this._workbookBinStart);

                for (const sheet of this.sheetList) {
                    const rId = `rId${sheet.sheetId}`;

                    const recLen = 4 + 12 + sheet.name.length * 2 + rId.length * 2;
                    const buf = Buffer.alloc(3 + recLen);
                    buf[0] = 156; buf[1] = 1; buf[2] = recLen;
                    let pos = 3;

                    buf.writeInt32LE(sheet.hidden ? 1 : 0, pos); pos += 4;
                    buf.writeInt32LE(sheet.sheetId, pos); pos += 4;

                    buf.writeInt32LE(rId.length, pos); pos += 4;
                    for (let i = 0; i < rId.length; i++) {
                        const c = rId.charCodeAt(i);
                        buf[pos++] = c & 0xFF; buf[pos++] = (c >> 8) & 0xFF;
                    }

                    buf.writeInt32LE(sheet.name.length, pos); pos += 4;
                    for (let i = 0; i < sheet.name.length; i++) {
                        const c = sheet.name.charCodeAt(i);
                        buf[pos++] = c & 0xFF; buf[pos++] = (c >> 8) & 0xFF;
                    }
                    wbBuffers.push(buf);
                }

                wbBuffers.push(this._workbookBinMiddle);

                if (this._autofilterIsOn) {
                    const filteredSheets = this.sheetList.filter(s => s.filterData);
                    const cnt = filteredSheets.length;

                    if (cnt > 0) {
                        wbBuffers.push(this._magicFilterExcel2016Fix0);

                        const firstByte = cnt <= 20
                            ? 0x10 + (cnt - 1) * 0x0C
                            : 0x80 + (cnt - 21) * 0x0C;
                        const countBuf = cnt <= 10
                            ? Buffer.from([firstByte, cnt, 0x00, 0x00, 0x00])
                             : Buffer.from([firstByte, Math.floor((cnt - 1) / 10), cnt, 0x00, 0x00, 0x00]);
                        wbBuffers.push(countBuf);

                        for (let nm = 0; nm < cnt; nm++) {
                            const sheetIndex = filteredSheets[nm].sheetId - 1;
                            const idxBuf = Buffer.alloc(12);
                            idxBuf.writeInt32LE(0, 0);
                            idxBuf[4] = sheetIndex;
                            idxBuf[8] = sheetIndex;
                            wbBuffers.push(idxBuf);
                        }

                        wbBuffers.push(Buffer.from([0xE2, 0x02, 0x00]));

                        for (let sheetNum = 0; sheetNum < cnt; sheetNum++) {
                            const sheet = filteredSheets[sheetNum];
                            this._writeFilterDefinedName(wbBuffers, sheet, sheetNum);
                        }
                    }
                }

                wbBuffers.push(this._workbookBinEnd);

                this.archive.append(Buffer.concat(wbBuffers), { name: 'xl/workbook.bin' });

                for (const sheet of this.sheetList) {
                    this.archive.append(this._binaryIndexBin, { name: `xl/worksheets/binaryIndex${sheet.sheetId}.bin` });
                }

                const hasComments = this.sheetList.some(sheet => (sheet.comments?.length ?? 0) > 0);
                let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`;
                if (hasComments) {
                    contentTypes += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
                }
                for (const sheet of this.sheetList) {
                    contentTypes += `<Override PartName="/${sheet.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>`;
                    contentTypes += `<Override PartName="/xl/worksheets/binaryIndex${sheet.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>`;
                    for (const part of sheet.tableParts ?? []) {
                        contentTypes += `<Override PartName="/${part.path}" ContentType="application/vnd.ms-excel.table"/>`;
                    }
                    if ((sheet.comments?.length ?? 0) > 0) {
                        contentTypes += `<Override PartName="/xl/comments/comment${sheet.sheetId}.bin" ContentType="application/vnd.ms-excel.comments"/>`;
                    }
                }
                contentTypes += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
                this.archive.append(contentTypes, { name: '[Content_Types].xml' });
                this.archive.append(XLSX_THEME_XML, { name: 'xl/theme/theme1.xml' });

                let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
                for (const sheet of this.sheetList) {
                    const rId = `rId${sheet.sheetId}`;
                    wbRels += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheet.nameInArchive}"/>`;
                }
                wbRels += `<Relationship Id="rId${this.sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this.sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
<Relationship Id="rId${this.sheetList.length + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;
                this.archive.append(wbRels, { name: 'xl/_rels/workbook.bin.rels' });

                const globalRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>`;
                this.archive.append(globalRels, { name: '_rels/.rels' });

                for (const sheet of this.sheetList) {
                    let wsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${sheet.sheetId}.bin"/>`;
                    for (const link of sheet.hyperlinks ?? []) {
                        if (!link.rId || link.spec.target === undefined) continue;
                        wsRels += `<Relationship Id="${link.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXmlText(link.spec.target)}" TargetMode="External"/>`;
                    }
                    for (const part of sheet.tableParts ?? []) {
                        wsRels += `<Relationship Id="${part.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/${part.path.split('/').pop()}"/>`;
                    }
                    if (sheet.commentRId !== undefined) {
                        wsRels += `<Relationship Id="${sheet.commentRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment${sheet.sheetId}.bin"/>`;
                    }
                    if (sheet.commentVmlRId !== undefined) {
                        wsRels += `<Relationship Id="${sheet.commentVmlRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/commentsDrawing${sheet.sheetId}.vml"/>`;
                    }
                    wsRels += '</Relationships>';
                    this.archive.append(wsRels, { name: `xl/worksheets/_rels/${sheet.nameInArchive}.rels` });
                }

                this.output.once('close', succeed);
                this.output.once('error', fail);
                this.archive.once('error', fail);
                this.archive.finalize();
            } catch (err) {
                fail(err);
            }
        });
    }
}

export default XlsbWriter;
