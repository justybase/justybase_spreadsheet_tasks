/**
 * BrowserXlsbWriter - Browser-compatible XLSB writer.
 * Port of XlsbWriter.ts using BrowserBigBuffer + BrowserZip instead of Node.js APIs.
 * Includes bold headers, frozen panes, and autofit columns.
 * @module BrowserXlsbWriter
 */
import { BrowserBigBuffer } from './BrowserBigBuffer';
import { BrowserZip } from './BrowserZip';
import {
    CellValue,
    ExcelError,
    FormulaCell,
    getCellStyle,
    getFormula,
    isRichText,
    unwrapCell,
} from '../Formats';
import {
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
    initColWidths,
    applyHeaderWidths,
    updateColWidthsFromRows,
    assertColumnWidth,
    defaultColWidth,
    timeToOaDate,
    dateHasTime,
    AUTOFIT_SAMPLE_ROWS,
    resolveAutofitRows,
} from '../writerHelpers';
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
} from '../writers/core/xlsbRecords';
import { XlsbStyleRegistry } from '../writers/core/xlsbStyles';
import { buildXlsbFormulaRecord } from '../writers/core/xlsbFormula';
import {
    buildXlsbHyperlinkRecord,
    buildXlsbMergeRecords,
    buildXlsbValidationRecords,
    parseXlsbRange,
} from '../writers/core/xlsbSheetFeatures';
import {
    DataValidationSpec,
    HyperlinkSpec,
    XlsxSheetFeatures,
} from '../writers/core/xlsxParts';
import { escapeXmlText } from '../xmlUtils';

const INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;

type BrowserHeaderStyle = 'bold' | 'fill' | 'bold+fill';

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

interface BrowserFilterData {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

interface BrowserSheetEntry {
    name: string;
    pathInArchive: string;
    hidden: boolean;
    nameInArchive: string;
    sheetId: number;
    filterHeaderRange: null;
    filterData?: BrowserFilterData;
    merges?: string[];
    dataValidations?: DataValidationSpec[];
    hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string; row: number; col: number }>;
}

export class BrowserXlsbWriter {
    private _zip: BrowserZip;
    private _sheetCount: number;
    private _sheetList: BrowserSheetEntry[];
    private _sstDic: Map<string, number>;
    private _sstCntUnique: number;
    private _sstCntAll: number;
    private _colWidths: number[];
    private _autofilterIsOn: boolean;
    private _oaEpoch: number;
    private _isStreaming: boolean;
    private _currentSheetBuffer: BrowserBigBuffer | null;
    private _currentSheetRowNum: number;
    private _currentSheetStartCol: number;
    private _currentSheetEndCol: number;
    private _currentSheetDoAutofilter: boolean;
    private _streamPending: BrowserBigBuffer | null;
    private _streamHeadFlushed: boolean;
    private _streamHasHeaders: boolean;
    private _streamDataRows: number;
    private _streamAutofitRows: number;
    private readonly _styles: XlsbStyleRegistry;
    private _sheet1Bytes: Uint8Array<ArrayBuffer>;
    private _workbookBinStart: Uint8Array<ArrayBuffer>;
    private _workbookBinMiddle: Uint8Array<ArrayBuffer>;
    private _workbookBinEnd: Uint8Array<ArrayBuffer>;
    private _stylesBin: Uint8Array<ArrayBuffer>;
    private _binaryIndexBin: Uint8Array<ArrayBuffer>;
    private _rRkIntegerLowerLimit: number;
    private _rRkIntegerUpperLimit: number;
    private _autoFilterStartBytes: Uint8Array<ArrayBuffer>;
    private _autoFilterEndBytes: Uint8Array<ArrayBuffer>;
    private _stickHeaderA1bytes: Uint8Array<ArrayBuffer>;
    private _magicFilterExcel2016Fix0: Uint8Array<ArrayBuffer>;
    private _magicFilterExcel2016Fix1: Uint8Array<ArrayBuffer>;
    private _magicFilterExcel2016Fix2: Uint8Array<ArrayBuffer>;
    private readonly _useSharedStrings: boolean;

    constructor(options: { useSharedStrings?: boolean } = {}) {
        this._useSharedStrings = options.useSharedStrings !== false;
        this._zip = new BrowserZip();
        this._sheetCount = 0;
        this._sheetList = [];
        this._sstDic = new Map();
        this._sstCntUnique = 0;
        this._sstCntAll = 0;
        this._colWidths = [];
        this._autofilterIsOn = false;
        this._oaEpoch = Date.UTC(1899, 11, 30);

        this._isStreaming = false;
        this._currentSheetBuffer = null;
        this._currentSheetRowNum = 0;
        this._currentSheetStartCol = 0;
        this._currentSheetEndCol = 0;
        this._currentSheetDoAutofilter = false;
        this._streamPending = null;
        this._streamHeadFlushed = true;
        this._streamHasHeaders = false;
        this._streamDataRows = 0;
        this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;

        this._styles = new XlsbStyleRegistry();

        this._sheet1Bytes = new Uint8Array([
            0x81, 0x01, 0x00, 0x93, 0x01, 0x17, 0xCB, 0x04,
            0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0x00, 0x00, 0x00, 0x00, 0x94, 0x01, 0x10,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x85, 0x01, 0x00, 0x89, 0x01, 0x1E, 0xDC, 0x03,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00,
            0x64, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x98, 0x01, 0x24, 0x03,
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
            0x00, 0x00, 0xDD, 0x03, 0x02, 0x10, 0x00, 0xDC,
            0x03, 0x30, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66,
            0xE6, 0x3F, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66,
            0xE6, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xE8, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xE8, 0x3F, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33,
            0xD3, 0x3F, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33,
            0xD3, 0x3F, 0x25, 0x06, 0x01, 0x00, 0x00, 0x10,
            0x00, 0x80, 0x80, 0x18, 0x10, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x82,
            0x01, 0x00
        ]);

        this._workbookBinStart = new Uint8Array([
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
            0x00, 0x00, 0x00, 0x78, 0x88, 0x01, 0x00, 0x8F, 0x01, 0x00
        ]);

        this._workbookBinMiddle = new Uint8Array([0x90, 0x01, 0x00]);

        this._workbookBinEnd = new Uint8Array([
            0x9D, 0x01, 0x1A, 0x35, 0xEA, 0x02, 0x00, 0x01, 0x00, 0x00, 0x00, 0x64, 0x00, 0x00, 0x00, 0xFC, 0xA9, 0xF1, 0xD2, 0x4D, 0x62, 0x50, 0x3F, 0x01,
            0x00, 0x00, 0x00, 0x6A, 0x00, 0x9B, 0x01, 0x01, 0x00, 0x23, 0x04, 0x03, 0x0F, 0x00, 0x00, 0xAB, 0x10, 0x01, 0x01, 0x24, 0x00, 0x84, 0x01, 0x00
        ]);

        this._stylesBin = new Uint8Array(XLSB_STYLES_TEMPLATE);

        this._binaryIndexBin = new Uint8Array([
            0x2A, 0x18, 0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x95,
            0x02, 0x00
        ]);

        this._rRkIntegerLowerLimit = -1 << 29;
        this._rRkIntegerUpperLimit = (1 << 29) - 1;

        this._autoFilterStartBytes = new Uint8Array([0xA1, 0x01, 0x10]);
        this._autoFilterEndBytes = new Uint8Array([0xA2, 0x01, 0x00]);

        this._stickHeaderA1bytes = new Uint8Array([
            0x97, 0x01, 0x1D, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0xF0, 0x3F, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x03
        ]);

        this._magicFilterExcel2016Fix0 = new Uint8Array([0xE1, 0x02, 0x00, 0xE5, 0x02, 0x00, 0xEA, 0x02]);
        this._magicFilterExcel2016Fix1 = new Uint8Array([
            0x27, 0x46, 0x21, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x0F, 0x00, 0x00, 0x00, 0x5F,
            0x00, 0x46, 0x00, 0x69, 0x00, 0x6C, 0x00, 0x74, 0x00, 0x65, 0x00, 0x72, 0x00, 0x44, 0x00, 0x61,
            0x00, 0x74, 0x00, 0x61, 0x00, 0x62, 0x00, 0x61, 0x00, 0x73, 0x00, 0x65, 0x00, 0x0F, 0x00, 0x00,
            0x00, 0x3B, 0xFF, 0x00
        ]);
        this._magicFilterExcel2016Fix2 = new Uint8Array([0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF]);
    }

    _sanitizeSheetName(name: string): string {
        if (!name || typeof name !== 'string') return `Sheet${this._sheetCount + 1}`;
        let s = name.replace(INVALID_SHEET_NAME_CHARS, '_');
        if (s.length > 31) s = s.substring(0, 31);
        if (s.trim().length === 0) s = `Sheet${this._sheetCount + 1}`;
        return s;
    }

    _buildStylesBin(): Uint8Array<ArrayBuffer> {
        return new Uint8Array(this._styles.toStylesBin(this._stylesBin));
    }

    /** Resolve column widths; explicit widths win over the conservative autofit. */
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

    /** Store supported sheet features; reject ones XLSB does not emit yet. */
    private _applySheetFeatures(sheet: BrowserSheetEntry, features: BrowserSheetOptions): void {
        if (features.conditionalFormats && features.conditionalFormats.length > 0) {
            throw new Error('Conditional formatting is not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.');
        }
        if (features.comments && features.comments.length > 0) {
            throw new Error('Cell comments are not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.');
        }
        if (features.merges) sheet.merges = features.merges;
        if (features.dataValidations) sheet.dataValidations = features.dataValidations;
        if (features.hyperlinks && features.hyperlinks.length > 0) {
            let nextRelId = 2;
            sheet.hyperlinks = features.hyperlinks.map(spec => {
                const range = parseXlsbRange(spec.ref);
                const internal = spec.target === undefined;
                return {
                    spec,
                    rId: internal ? undefined : `rId${nextRelId++}`,
                    row: range.r1,
                    col: range.c1,
                };
            });
        }
    }

    private _writeFeatureRecords(bb: BrowserBigBuffer, sheet: BrowserSheetEntry): void {
        if (sheet.merges && sheet.merges.length > 0) {
            bb.write(buildXlsbMergeRecords(sheet.merges));
        }
        if (sheet.dataValidations && sheet.dataValidations.length > 0) {
            bb.write(buildXlsbValidationRecords(sheet.dataValidations));
        }
        for (const link of sheet.hyperlinks ?? []) {
            bb.write(buildXlsbHyperlinkRecord(link.row, link.col, link.spec, link.rId));
        }
    }

    addSheet(sheetName: string, hidden = false): void {
        const s = this._sanitizeSheetName(sheetName);
        this._sheetCount++;
        this._sheetList.push({
            name: s,
            pathInArchive: `xl/worksheets/sheet${this._sheetCount}.bin`,
            hidden,
            nameInArchive: `sheet${this._sheetCount}.bin`,
            sheetId: this._sheetCount,
            filterHeaderRange: null
        });
    }

    startSheet(sheetName: string, columnCount: number, headers?: string[], options: BrowserStartSheetOptions = {}): void {
        if (this._isStreaming) throw new Error('Already streaming');
        assertColumnCount(columnCount);
        if (headers !== undefined) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        const { hidden = false, doAutofilter = true, headerStyle = 'bold' } = options;
        let hs = 3; if (headerStyle === 'fill') hs = 4; else if (headerStyle === 'bold+fill') hs = 5;
        this.addSheet(sheetName, hidden);
        const sheet = this._sheetList[this._sheetCount - 1];
        this._applySheetFeatures(sheet, options);
        this._isStreaming = true;
        this._currentSheetBuffer = new BrowserBigBuffer();
        this._currentSheetRowNum = 0;
        this._currentSheetStartCol = 0;
        this._currentSheetEndCol = columnCount;
        this._currentSheetDoAutofilter = doAutofilter && headers !== undefined;
        this._colWidths = this._resolveColumnWidths(
            columnCount,
            headers,
            undefined,
            options,
            doAutofilter && headers !== undefined,
        );

        // Delay the head until the first AUTOFIT_SAMPLE_ROWS rows are measured.
        this._streamPending = new BrowserBigBuffer();
        this._streamHeadFlushed = false;
        this._streamHasHeaders = headers !== undefined;
        this._streamDataRows = 0;
        this._streamAutofitRows = resolveAutofitRows(options.autofitRows);

        if (headers) {
            const pending = this._streamPending;
            this.createRowHeader(pending, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
            for (let c = 0; c < headers.length; c++) this.writeString(pending, headers[c], c, hs);
            this._currentSheetRowNum++;
        }
        if (this._streamAutofitRows === 0) this._flushStreamHead();
    }

    private _flushStreamHead(): void {
        if (this._streamHeadFlushed) return;
        const bb = this._currentSheetBuffer;
        if (!this._isStreaming || bb === null) throw new Error('Not streaming');
        this._writeSheetHeader(bb, this._currentSheetStartCol, this._currentSheetEndCol, this._currentSheetDoAutofilter);
        const pending = this._streamPending;
        if (pending) {
            for (const chunk of pending.getChunks()) bb.write(chunk);
        }
        this._streamPending = null;
        this._streamHeadFlushed = true;
    }

    private _writeStreamRow(bb: BrowserBigBuffer, row: CellValue[]): void {
        assertRowCount(this._currentSheetRowNum + 1);
        this.createRowHeader(bb, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
        for (let c = 0; c < row.length; c++) this._writeCell(bb, row[c], c);
        this._currentSheetRowNum++;
    }

    writeRow(row: CellValue[]): void {
        const bb = this._currentSheetBuffer;
        if (!this._isStreaming || bb === null) throw new Error('Not streaming');
        if (row.length !== (this._currentSheetEndCol - this._currentSheetStartCol)) {
            throw new Error(`Row length mismatch. Expected ${this._currentSheetEndCol - this._currentSheetStartCol} columns, got ${row.length}`);
        }
        if (!this._streamHeadFlushed && this._streamPending) {
            this._streamDataRows++;
            updateColWidthsFromRows(this._colWidths, [row]);
            this._writeStreamRow(this._streamPending, row);
            if (this._streamDataRows >= this._streamAutofitRows) this._flushStreamHead();
            return;
        }
        this._writeStreamRow(bb, row);
    }

    endSheet(): void {
        this._flushStreamHead();
        const bb = this._currentSheetBuffer;
        if (!this._isStreaming || bb === null) throw new Error('Not streaming');
        bb.write(this._sheet1Bytes.subarray(218, 290));
        
        if (this._currentSheetDoAutofilter) {
            this._autofilterIsOn = true;
            bb.write(this._autoFilterStartBytes);
            const rBuf = new Uint8Array(8);
            const rView = new DataView(rBuf.buffer);
            rView.setInt32(0, 0, true);
            rView.setInt32(4, this._currentSheetRowNum - 1, true);
            bb.write(rBuf);
            const cBuf = new Uint8Array(8);
            const cView = new DataView(cBuf.buffer);
            cView.setInt32(0, this._currentSheetStartCol, true);
            cView.setInt32(4, this._currentSheetEndCol - 1, true);
            bb.write(cBuf);
            bb.write(this._autoFilterEndBytes);
            
            this._sheetList[this._sheetCount - 1].filterData = {
                startRow: 0, endRow: this._currentSheetRowNum - 1,
                startColumn: this._currentSheetStartCol, endColumn: this._currentSheetEndCol - 1
            };
        }

        this._writeFeatureRecords(bb, this._sheetList[this._sheetCount - 1]);

        bb.write(this._sheet1Bytes.subarray(290));
        bb.patchInt32LE(36, Math.max(0, this._currentSheetRowNum - 1));
        this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
        this._isStreaming = false;
        this._currentSheetBuffer = null;
        this._streamPending = null;
        this._streamHeadFlushed = true;
        this._streamHasHeaders = false;
        this._streamDataRows = 0;
        this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    }

    writeSheet(rows: CellValue[][], headers: string[] | null = null, options: BrowserSheetOptions = {}): void {
        const doAutofilter = options.doAutofilter !== false;
        const headerStyle = options.headerStyle || 'bold';
        let hs = 3; if (headerStyle === 'fill') hs = 4; else if (headerStyle === 'bold+fill') hs = 5;
        const bb = new BrowserBigBuffer();
        const colCount = rows.length > 0 ? rows[0].length : (headers ? headers.length : 0);
        if (colCount > 0) assertColumnCount(colCount);
        const totalBr = rows.length + (headers ? 1 : 0);
        assertRowCount(totalBr);
        if (headers) {
            assertRowFitsColumns('header', headers.length, colCount);
            for (const h of headers) assertCellTextLength(h);
        }
        for (let r = 0; r < rows.length; r++) assertRowFitsColumns('row', rows[r].length, colCount);
        const sheet = this._sheetList[this._sheetCount - 1];
        if (!sheet) throw new Error('Call addSheet() before writeSheet()');
        this._applySheetFeatures(sheet, options);
        this._colWidths = this._resolveColumnWidths(
            colCount,
            headers,
            rows,
            options,
            doAutofilter && headers !== null,
        );
        
        const doAf = doAutofilter && !!headers;
        this._writeSheetHeader(bb, 0, colCount, doAf);
        
        let rn = 0;
        if (headers) {
            this.createRowHeader(bb, rn, 0, colCount);
            for (let c = 0; c < headers.length; c++) this.writeString(bb, headers[c], c, hs);
            rn++;
        }
        
        for (let r = 0; r < rows.length; r++) {
            this.createRowHeader(bb, rn, 0, colCount);
            for (let c = 0; c < rows[r].length; c++) this._writeCell(bb, rows[r][c], c);
            rn++;
        }
        
        bb.write(this._sheet1Bytes.subarray(218, 290));
        
        if (doAf) {
            this._autofilterIsOn = true;
            bb.write(this._autoFilterStartBytes);
            const rBuf = new Uint8Array(8), rView = new DataView(rBuf.buffer);
            rView.setInt32(0, 0, true); rView.setInt32(4, rn - 1, true); bb.write(rBuf);
            const cBuf = new Uint8Array(8), cView = new DataView(cBuf.buffer);
            cView.setInt32(0, 0, true); cView.setInt32(4, colCount - 1, true); bb.write(cBuf);
            bb.write(this._autoFilterEndBytes);
            this._sheetList[this._sheetCount - 1].filterData = {
                startRow: 0, endRow: rn - 1, startColumn: 0, endColumn: colCount - 1
            };
        }

        this._writeFeatureRecords(bb, sheet);

        bb.write(this._sheet1Bytes.subarray(290));
        bb.patchInt32LE(36, Math.max(0, rn - 1));
        this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
    }

    _writeSheetHeader(bb: BrowserBigBuffer, startCol: number, endCol: number, doAf: boolean): void {
        const sh = new Uint8Array(this._sheet1Bytes);
        const shView = new DataView(sh.buffer);
        shView.setInt32(40, startCol, true);
        shView.setInt32(44, endCol - 1, true);
        if (this._sheetCount !== 1) sh[54] = 0x9C;
        
        bb.write(sh.subarray(0, 84));
        if (doAf) bb.write(this._stickHeaderA1bytes);
        bb.write(sh.subarray(84, 159));
        
        bb.writeByte(134); bb.writeByte(3); bb.writeByte(0);
        for (let i = startCol; i < endCol; i++) {
            bb.writeByte(60); bb.writeByte(18);
            bb.writeInt32LE(i); bb.writeInt32LE(i);
            const w = defaultColWidth(this._colWidths[i]);
            bb.writeInt32LE(Math.max(0, Math.min(255 * 256, Math.round(w * 256))));
            bb.writeInt32LE(0);
            bb.writeByte(2); bb.writeByte(0);
        }
        bb.writeByte(135); bb.writeByte(3); bb.writeByte(0);
        bb.write(sh.subarray(159, 175));
        bb.write(new Uint8Array([38, 0]));
    }

    _writeCell(bb: BrowserBigBuffer, raw: CellValue, c: number): void {
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
            bb.write(buildXlsbFormulaRecord(c, styleNum, formula, cached));
            return;
        }

        const val = unwrapCell(raw);
        if (val === null || val === undefined) return;
        if (isRichText(val)) {
            throw new Error('Rich text is not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.');
        }

        if (typeof val === 'number') {
            if (Number.isInteger(val) && val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
                this.writeRkNumberInteger(bb, val, c, styleNum);
            } else {
                this.writeDouble(bb, val, c, styleNum);
            }
        } else if (val instanceof ExcelError) {
            this.writeError(bb, val.code, c, styleNum);
        } else if (typeof val === 'bigint') {
            this.writeString(bb, val.toString(), c, 0, styleNum);
        } else if (typeof val === 'boolean') {
            this.writeBool(bb, val, c, styleNum);
        } else if (val instanceof Date) {
            const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
            const dateStyle = dateHasTime(val) ? 2 : 1;
            this.writeDouble(bb, oaDate, c, styleNum !== 0 ? styleNum : dateStyle);
        } else {
            this.writeString(bb, val.toString(), c, 0, styleNum);
        }
    }

    createRowHeader(bb: BrowserBigBuffer, rn: number, sc: number, ec: number): void {
        writeXlsbRowHeader(bb, rn, sc, ec);
    }
    writeRkNumberInteger(bb: BrowserBigBuffer, val: number, colNum: number, styleNum = 0): void {
        writeXlsbRkInteger(bb, val, colNum, styleNum);
    }
    writeDouble(bb: BrowserBigBuffer, val: number, colNum: number, styleNum = 0): void {
        writeXlsbDouble(bb, val, colNum, styleNum);
    }
    writeBool(bb: BrowserBigBuffer, val: boolean, colNum: number, styleNum = 0): void {
        writeXlsbBool(bb, val, colNum, styleNum);
    }
    writeError(bb: BrowserBigBuffer, code: string, colNum: number, styleNum = 0): void {
        writeXlsbError(bb, code, colNum, styleNum);
    }
    writeDateTime(bb: BrowserBigBuffer, date: Date, colNum: number): void {
        this.writeDouble(bb, timeToOaDate(date.getTime(), this._oaEpoch), colNum, dateHasTime(date) ? 2 : 1);
    }
    writeString(bb: BrowserBigBuffer, val: string, colNum: number, styleIndex = 0, styleOverride?: number): void {
        assertCellTextLength(val);
        const finalStyle = styleOverride !== undefined ? styleOverride : styleIndex;
        if (!this._useSharedStrings) {
            writeXlsbInlineString(bb, val, colNum, finalStyle);
            return;
        }
        let index = this._sstDic.get(val);
        if (index === undefined) {
            index = this._sstCntUnique++;
            this._sstDic.set(val, index);
        }
        this._sstCntAll++;
        writeXlsbIsst(bb, index, colNum, finalStyle);
    }

    _saveSst(): void {
        const bb = new BrowserBigBuffer();
        writeXlsbSharedStrings(bb, this._sstDic.keys(), this._sstCntAll, this._sstCntUnique);
        this._zip.addFile('xl/sharedStrings.bin', bb.toUint8Array());
    }

    _writeFilterDefinedName(wbBuffers: Uint8Array<ArrayBuffer>[], sheet: BrowserSheetEntry & { filterData: BrowserFilterData }, sheetNum: number): void {
        const fd = sheet.filterData;
        const fix1 = new Uint8Array(this._magicFilterExcel2016Fix1);
        fix1[7] = sheet.sheetId - 1;
        fix1[fix1.length - 2] = sheetNum;
        wbBuffers.push(fix1);
        
        const rBuf = new Uint8Array(8), rView = new DataView(rBuf.buffer);
        rView.setInt32(0, fd.startRow, true); rView.setInt32(4, fd.endRow, true); wbBuffers.push(rBuf);
        const cBuf = new Uint8Array(4), cView = new DataView(cBuf.buffer);
        cView.setInt16(0, fd.startColumn, true); cView.setInt16(2, fd.endColumn, true); wbBuffers.push(cBuf);
        wbBuffers.push(this._magicFilterExcel2016Fix2);
    }

    finalize(): Blob {
        this._saveSst();
        this._zip.addFile('xl/styles.bin', this._buildStylesBin());
        
        const wbBuffers = [this._workbookBinStart];
        for (const sheet of this._sheetList) {
            const rId = `rId${sheet.sheetId}`;
            const sn = sheet.name;
            const rl = 4 + 12 + sn.length * 2 + rId.length * 2;
            const buf = new Uint8Array(3 + rl);
            const view = new DataView(buf.buffer);
            buf[0] = 156; buf[1] = 1; buf[2] = rl;
            let pos = 3;
            view.setInt32(pos, sheet.hidden ? 1 : 0, true); pos += 4;
            view.setInt32(pos, sheet.sheetId, true); pos += 4;
            view.setInt32(pos, rId.length, true); pos += 4;
            for (let i = 0; i < rId.length; i++) { const c = rId.charCodeAt(i); buf[pos++] = c & 0xFF; buf[pos++] = (c >> 8) & 0xFF; }
            view.setInt32(pos, sn.length, true); pos += 4;
            for (let i = 0; i < sn.length; i++) { const c = sn.charCodeAt(i); buf[pos++] = c & 0xFF; buf[pos++] = (c >> 8) & 0xFF; }
            wbBuffers.push(buf);
        }
        
        wbBuffers.push(this._workbookBinMiddle);
        
        if (this._autofilterIsOn) {
            const filtered = this._sheetList.filter((s): s is BrowserSheetEntry & { filterData: BrowserFilterData } => s.filterData !== undefined);
            if (filtered.length > 0) {
                wbBuffers.push(this._magicFilterExcel2016Fix0);
                const cnt = filtered.length;
                const firstByte = cnt <= 20 ? 0x10 + (cnt - 1) * 0x0C : 0x80 + (cnt - 21) * 0x0C;
                wbBuffers.push(cnt <= 10
                    ? new Uint8Array([firstByte, cnt, 0, 0, 0])
                    : new Uint8Array([firstByte, Math.floor((cnt - 1) / 10), cnt, 0, 0, 0]));
                for (let i = 0; i < filtered.length; i++) {
                    const si = filtered[i].sheetId - 1;
                    const ib = new Uint8Array(12); ib[4] = si; ib[8] = si; wbBuffers.push(ib);
                }
                wbBuffers.push(new Uint8Array([0xE2, 0x02, 0x00]));
                for (let i = 0; i < filtered.length; i++) this._writeFilterDefinedName(wbBuffers, filtered[i], i);
            }
        }
        
        wbBuffers.push(this._workbookBinEnd);
        
        let totalWbLen = 0; for (const b of wbBuffers) totalWbLen += b.length;
        const wbFinal = new Uint8Array(totalWbLen); let wbOffset = 0;
        for (const b of wbBuffers) { wbFinal.set(b, wbOffset); wbOffset += b.length; }
        this._zip.addFile('xl/workbook.bin', wbFinal);
        
        for (const s of this._sheetList) this._zip.addFile(`xl/worksheets/binaryIndex${s.sheetId}.bin`, this._binaryIndexBin);
        
        let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`;
        for (const s of this._sheetList) {
            ct += `<Override PartName="/${s.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>`;
            ct += `<Override PartName="/xl/worksheets/binaryIndex${s.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>`;
        }
        ct += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>`;
        this._zip.addFile('[Content_Types].xml', ct);

        let wr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        for (const s of this._sheetList) wr += `<Relationship Id="rId${s.sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.nameInArchive}"/>`;
        wr += `<Relationship Id="rId${this._sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this._sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
</Relationships>`;
        this._zip.addFile('xl/_rels/workbook.bin.rels', wr);

        this._zip.addFile('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>`);

        for (const s of this._sheetList) {
            let ws = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${s.sheetId}.bin"/>`;
            for (const link of s.hyperlinks ?? []) {
                if (!link.rId || link.spec.target === undefined) continue;
                ws += `<Relationship Id="${link.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXmlText(link.spec.target)}" TargetMode="External"/>`;
            }
            ws += '</Relationships>';
            this._zip.addFile(`xl/worksheets/_rels/${s.nameInArchive}.rels`, ws);
        }
        
        return this._zip.toBlob();
    }
}
export default BrowserXlsbWriter;
