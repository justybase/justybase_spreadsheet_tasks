import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import archiver from 'archiver';
import { Readable } from 'stream';
import { BigBuffer } from './BigBuffer';
import {
    CellValue,
    ExcelError,
    FormulaCell,
    getCellStyle,
    getFormula,
    isRichText,
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
    XLSB_STYLES_TEMPLATE,
} from './writers/core/xlsbRecords';
import { XlsbStyleRegistry } from './writers/core/xlsbStyles';
import { buildXlsbFormulaRecord } from './writers/core/xlsbFormula';
import {
    buildXlsbHyperlinkRecord,
    buildXlsbMergeRecords,
    buildXlsbValidationRecords,
    parseXlsbRange,
} from './writers/core/xlsbSheetFeatures';
import {
    DataValidationSpec,
    HyperlinkSpec,
    XlsxSheetFeatures,
} from './writers/core/xlsxParts';
import { escapeXmlText } from './xmlUtils';

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
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string; row: number; col: number }>;
}

/** Options accepted by the XLSB streaming and batch sheet APIs. */
export interface XlsbSheetOptions extends XlsxSheetFeatures {
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
    private sstCntUnique: number = 0;
    private sstCntAll: number = 0;
    private _autofilterIsOn: boolean = false;

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

    /** Store supported sheet features; reject the ones XLSB does not emit yet. */
    private _applySheetFeatures(sheet: SheetInfo, features: XlsbSheetOptions): void {
        if (features.conditionalFormats && features.conditionalFormats.length > 0) {
            throw new Error('Conditional formatting is not supported by XlsbWriter yet; use XlsxWriter.');
        }
        if (features.comments && features.comments.length > 0) {
            throw new Error('Cell comments are not supported by XlsbWriter yet; use XlsxWriter.');
        }
        if (features.merges) sheet.merges = features.merges;
        if (features.dataValidations) sheet.dataValidations = features.dataValidations;
        if (features.hyperlinks && features.hyperlinks.length > 0) {
            let nextRelId = 2; // rId1 is the worksheet binary index
            sheet.hyperlinks = features.hyperlinks.map(spec => {
                const range = parseXlsbRange(spec.ref);
                const internal = spec.target === undefined;
                const rId = internal ? undefined : `rId${nextRelId++}`;
                return { spec, rId, row: range.r1, col: range.c1 };
            });
        }
    }

    /** Merged cells, validations and hyperlinks go after the autoFilter. */
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
    }

    /** Write one cell using the current style registry and formula compiler. */
    private _writeCell(bigBuf: BigBuffer, raw: CellValue, col: number): void {
        if (raw === null || raw === undefined) return;

        const style = getCellStyle(raw);
        let styleNum = style !== null ? this._styles.register(style) : 0;

        const formula = getFormula(raw);
        if (formula !== null) {
            const rawCached = (raw as FormulaCell).value;
            let cached: string | number | boolean | ExcelError | null | undefined;
            if (rawCached instanceof Date) {
                cached = timeToOaDate(rawCached.getTime(), this._oaEpoch);
                if (styleNum === 0) styleNum = dateHasTime(rawCached) ? 2 : 1;
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
            throw new Error('Rich text is not supported by XlsbWriter yet; use XlsxWriter.');
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
            const dateStyle = dateHasTime(val) ? 2 : 1;
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
        this._applySheetFeatures(sheet, options);

        this.stream.begin(columnCount, doAutofilter && headers !== undefined, new BigBuffer());

        this.stream.colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            undefined,
            options,
            doAutofilter && headers !== undefined,
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
            this.createRowHeader(pending, this.stream.rowNum, this.stream.startCol, this.stream.endCol);
            for (let c = 0; c < headers.length; c++) {
                this.writeString(pending, headers[c], c, true);
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

        const sheetHeader = Buffer.from(this._sheet1Bytes);
        sheetHeader.writeInt32LE(this.stream.startCol, 40);
        sheetHeader.writeInt32LE(this.stream.endCol - 1, 44);

        if (this.sheetCount !== 1) {
            sheetHeader[54] = 0x9C;
        }

        bigBuf.write(sheetHeader.subarray(0, 84));

        // Write sticky header if headers are present
        if (this._streamHasHeaders) {
            bigBuf.write(this._stickHeaderA1bytes);
        }

        bigBuf.write(sheetHeader.subarray(84, 159));

        // Write column definitions
        bigBuf.writeByte(134);
        bigBuf.writeByte(3);
        bigBuf.writeByte(0);

        for (let i = this.stream.startCol; i < this.stream.endCol; i++) {
            bigBuf.writeByte(60);
            bigBuf.writeByte(18);

            bigBuf.writeInt32LE(i);
            bigBuf.writeInt32LE(i);

            const width = defaultColWidth(this.stream.colWidths[i]);
            bigBuf.writeInt32LE(Math.max(0, Math.min(255 * 256, Math.round(width * 256))));

            bigBuf.writeInt32LE(0);
            bigBuf.writeByte(2);
            bigBuf.writeByte(0);
        }

        bigBuf.writeByte(135);
        bigBuf.writeByte(3);
        bigBuf.writeByte(0);

        bigBuf.write(sheetHeader.subarray(159, 175));
        bigBuf.write(Buffer.from([38, 0]));

        if (pending) {
            for (const chunk of pending.getChunks()) bigBuf.write(chunk);
        }
        this._streamPending = null;
        this._streamHeadFlushed = true;
    }

    private _writeStreamRow(bigBuf: BigBuffer, row: CellValue[]): void {
        assertRowCount(this.stream.rowNum + 1);
        this.createRowHeader(bigBuf, this.stream.rowNum, this.stream.startCol, this.stream.endCol);
        for (let c = 0; c < row.length; c++) {
            this._writeCell(bigBuf, row[c], c);
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
        const sheetHeader = Buffer.from(this._sheet1Bytes);

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
            const sheet = this.sheetList[this.sheetCount - 1];
            sheet.filterData = {
                startRow: 0,
                endRow: endRow - 1,
                startColumn: this.stream.startCol,
                endColumn: this.stream.endCol - 1
            };
        }

        this._writeFeatureRecords(bigBuf, this.sheetList[this.sheetCount - 1]);

        bigBuf.write(sheetHeader.subarray(290));

        // BrtWsDim was written with row 0; patch the inclusive last row now
        // that the row count is known (fixed offset 36 in the sheet header).
        bigBuf.patchInt32LE(36, Math.max(0, this.stream.rowNum - 1));

        // Add to archive
        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: this.sheetList[this.sheetCount - 1].pathInArchive
        });

        // Reset streaming state
        this.stream.end();
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
        this._applySheetFeatures(sheet, options);

        const colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            rows,
            options,
            doAutofilter && headers !== null && headers !== undefined,
        );

        const sheetHeader = Buffer.from(this._sheet1Bytes);
        const startCol = 0;
        const endCol = columnCount;

        sheetHeader.writeInt32LE(startCol, 40);
        sheetHeader.writeInt32LE(endCol - 1, 44);

        if (this.sheetCount !== 1) {
            sheetHeader[54] = 0x9C;
        }

        bigBuf.write(sheetHeader.subarray(0, 84));

        if (headers) {
            bigBuf.write(this._stickHeaderA1bytes);
        }

        bigBuf.write(sheetHeader.subarray(84, 159));

        bigBuf.writeByte(134);
        bigBuf.writeByte(3);
        bigBuf.writeByte(0);

        for (let i = startCol; i < endCol; i++) {
            bigBuf.writeByte(60);
            bigBuf.writeByte(18);

            bigBuf.writeInt32LE(i);
            bigBuf.writeInt32LE(i);

            const width = defaultColWidth(colWidths[i]);
            bigBuf.writeInt32LE(Math.max(0, Math.min(255 * 256, Math.round(width * 256))));

            bigBuf.writeInt32LE(0);
            bigBuf.writeByte(2);
            bigBuf.writeByte(0);
        }

        bigBuf.writeByte(135);
        bigBuf.writeByte(3);
        bigBuf.writeByte(0);

        bigBuf.write(sheetHeader.subarray(159, 175));
        bigBuf.write(Buffer.from([38, 0]));

        let rowNum = 0;

        if (headers) {
            this.createRowHeader(bigBuf, rowNum, startCol, endCol);
            for (let c = 0; c < headers.length; c++) {
                this.writeString(bigBuf, headers[c], c, true);
            }
            rowNum++;
        }

        for (let r = 0; r < rows.length; r++) {
            this.createRowHeader(bigBuf, rowNum, startCol, endCol);
            const row = rows[r];
            for (let c = 0; c < row.length; c++) {
                this._writeCell(bigBuf, row[c], c);
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

            const sheet = this.sheetList[this.sheetCount - 1];
            sheet.filterData = {
                startRow: 0,
                endRow: rows.length,
                startColumn: startCol,
                endColumn: endCol - 1
            };
        }

        this._writeFeatureRecords(bigBuf, sheet);

        bigBuf.write(sheetHeader.subarray(290));

        bigBuf.patchInt32LE(36, Math.max(0, rowNum - 1));

        this.archive.append(Readable.from(bigBuf.getChunks()), { name: this.sheetList[this.sheetCount - 1].pathInArchive });
    }



    createRowHeader(bigBuf: BigBuffer, rowNumber: number, startCol: number, endCol: number): void {
        writeXlsbRowHeader(bigBuf, rowNumber, startCol, endCol);
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
        this.writeDouble(bigBuf, oaDate, colNum, dateHasTime(date) ? 2 : 1);
    }

    writeString(bigBuf: BigBuffer, val: string, colNum: number, bolded: boolean = false, styleOverride?: number): void {
        assertCellTextLength(val);
        const styleNum = styleOverride !== undefined ? styleOverride : (bolded ? 3 : 0);

        if (!this._useSharedStrings) {
            writeXlsbInlineString(bigBuf, val, colNum, styleNum);
            return;
        }

        let index: number;
        if (this.sstDic.has(val)) {
            index = this.sstDic.get(val)!;
        } else {
            index = this.sstCntUnique;
            this.sstDic.set(val, index);
            this.sstCntUnique++;
        }
        this.sstCntAll++;

        writeXlsbIsst(bigBuf, index, colNum, styleNum);
    }

    private saveSst(): void {
        const bigBuf = new BigBuffer();
        writeXlsbSharedStrings(bigBuf, this.sstDic.keys(), this.sstCntAll, this.sstCntUnique);
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

                this.archive.append(Buffer.from(this._buildStylesBin()), { name: 'xl/styles.bin' });

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

                let contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`;
                for (const sheet of this.sheetList) {
                    contentTypes += `<Override PartName="/${sheet.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>`;
                    contentTypes += `<Override PartName="/xl/worksheets/binaryIndex${sheet.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>`;
                }
                contentTypes += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>`;
                this.archive.append(contentTypes, { name: '[Content_Types].xml' });

                let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
                for (const sheet of this.sheetList) {
                    const rId = `rId${sheet.sheetId}`;
                    wbRels += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheet.nameInArchive}"/>`;
                }
                wbRels += `<Relationship Id="rId${this.sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this.sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
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
