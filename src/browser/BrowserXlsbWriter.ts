/**
 * BrowserXlsbWriter - Browser-compatible XLSB writer.
 * Port of XlsbWriter.ts using BrowserBigBuffer + BrowserZip instead of Node.js APIs.
 * Includes bold headers, frozen panes, and autofit columns.
 * @module BrowserXlsbWriter
 */
import { BrowserBigBuffer } from './BrowserBigBuffer';
import { BrowserZip } from './BrowserZip';
import { CellValue, isFormattedCell, unwrapCell, getFormat } from '../Formats';
import {
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
    timeToOaDate,
} from '../writerHelpers';

const INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;

type BrowserHeaderStyle = 'bold' | 'fill' | 'bold+fill';

interface BrowserSheetOptions {
    doAutofilter?: boolean;
    headerStyle?: BrowserHeaderStyle;
}

interface BrowserStartSheetOptions extends BrowserSheetOptions {
    hidden?: boolean;
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
    private _fmtMap: Map<string, { nid: number; xf: number }>;
    private _nextNid: number;
    private _nextXf: number;
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

    constructor() {
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

        this._fmtMap = new Map();
        this._nextNid = 167;
        this._nextXf = 4;

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

        this._stylesBin = new Uint8Array([
            0x96, 0x02, 0x00, 0xE7, 0x04, 0x04, 0x02,
            0x00, 0x00, 0x00, 0x2C, 0x2C, 0xA4, 0x00, 0x13,
            0x00, 0x00, 0x00, 0x79, 0x00, 0x79, 0x00, 0x79,
            0x00, 0x79, 0x00, 0x5C, 0x00, 0x2D, 0x00, 0x6D,
            0x00, 0x6D, 0x00, 0x5C, 0x00, 0x2D, 0x00, 0x64,
            0x00, 0x64, 0x00, 0x5C, 0x00, 0x20, 0x00, 0x68,
            0x00, 0x68, 0x00, 0x3A, 0x00, 0x6D, 0x00, 0x6D,
            0x00, 0x2C, 0x1E, 0xA6, 0x00, 0x0C, 0x00, 0x00,
            0x00, 0x79, 0x00, 0x79, 0x00, 0x79, 0x00, 0x79,
            0x00, 0x5C, 0x00, 0x2D, 0x00, 0x6D, 0x00, 0x6D,
            0x00, 0x5C, 0x00, 0x2D, 0x00, 0x64, 0x00, 0x64,
            0x00, 0xE8, 0x04, 0x00, 0xE3, 0x04, 0x04, 0x02,
            0x00, 0x00, 0x00,
            0x2B, 0x27, 0xDC, 0x00, 0x00, 0x00, 0x90, 0x01,
            0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x07, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x02, 0x07,
            0x00, 0x00, 0x00, 0x43, 0x00, 0x61, 0x00, 0x6C,
            0x00, 0x69, 0x00, 0x62, 0x00, 0x72, 0x00, 0x69,
            0x00,
            0x2B, 0x27, 0xDC, 0x00, 0x01, 0x00, 0xBC, 0x02,
            0x00, 0x00, 0x00, 0x02, 0xEE, 0x00, 0x07, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x02, 0x07,
            0x00, 0x00, 0x00, 0x43, 0x00, 0x61, 0x00, 0x6C,
            0x00, 0x69, 0x00, 0x62, 0x00, 0x72, 0x00, 0x69,
            0x00,
            0x25, 0x06, 0x01, 0x00, 0x02, 0x0E, 0x00, 0x80, 0x81, 0x08, 0x00, 0x26, 0x00, 0xE4, 0x04, 0x00, 0xDB, 0x04, 0x04, 0x02, 0x00, 0x00, 0x00,
            0x2D, 0x44, 0x00, 0x00, 0x00, 0x00, 0x03, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x03, 0x41, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x2D, 0x44, 0x11, 0x00, 0x00, 0x00, 0x03, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x03, 0x41, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0xDC, 0x04, 0x00, 0xE5, 0x04, 0x04, 0x01, 0x00, 0x00, 0x00, 0x2E, 0x33, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE6, 0x04, 0x00, 0xF2,
            0x04, 0x04, 0x01, 0x00, 0x00, 0x00, 0x2F, 0x10, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00,
            0x00, 0xF3, 0x04, 0x00, 0xE9, 0x04, 0x04,
            0x06, // 6 xfs
            0x00, 0x00, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0xA4, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0xA6, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00, // index=3: bold
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00, // index=4: fill
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00, // index=5: bold+fill
            0xEA, 0x04, 0x00, 0xEB, 0x04, 0x04, 0x01, 0x00,
            0x00, 0x00, 0x25, 0x06, 0x01, 0x00, 0x02, 0x11, 0x00, 0x80, 0x80, 0x18, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x30, 0x1C, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x4E,
            0x00, 0x6F, 0x00, 0x72, 0x00, 0x6D, 0x00, 0x61, 0x00, 0x6C, 0x00, 0x6E, 0x00, 0x79, 0x00, 0xEC, 0x04, 0x00, 0xF9, 0x03, 0x04, 0x00, 0x00,
            0x00, 0x00, 0xFA, 0x03, 0x00, 0xFC, 0x03, 0x50, 0x00, 0x00, 0x00, 0x00, 0x11, 0x00, 0x00, 0x00, 0x54, 0x00, 0x61, 0x00, 0x62, 0x00, 0x6C,
            0x00, 0x65, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00, 0x6C, 0x00, 0x65, 0x00, 0x4D, 0x00, 0x65, 0x00, 0x64, 0x00, 0x69, 0x00, 0x75, 0x00,
            0x6D, 0x00, 0x32, 0x00, 0x11, 0x00, 0x00, 0x00, 0x50, 0x00, 0x69, 0x00, 0x76, 0x00, 0x6F, 0x00, 0x74, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79,
            0x00, 0x6C, 0x00, 0x65, 0x00, 0x4C, 0x00, 0x69, 0x00, 0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0x36, 0x00, 0xFD, 0x03, 0x00, 0x23,
            0x04, 0x02, 0x0E, 0x00, 0x00, 0xEB, 0x08, 0x00, 0xF6, 0x08, 0x2A, 0x00, 0x00, 0x00, 0x00, 0x11, 0x00, 0x00, 0x00, 0x53, 0x00, 0x6C, 0x00,
            0x69, 0x00, 0x63, 0x00, 0x65, 0x00, 0x72, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00, 0x6C, 0x00, 0x65, 0x00, 0x4C, 0x00, 0x69, 0x00, 0x67,
            0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0xF7, 0x08, 0x00, 0xEC, 0x08, 0x00, 0x24, 0x00, 0x23, 0x04, 0x03, 0x0F, 0x00, 0x00, 0xB0, 0x10,
            0x00, 0xB2, 0x10, 0x32, 0x00, 0x00, 0x00, 0x00, 0x15, 0x00, 0x00, 0x00, 0x54, 0x00, 0x69, 0x00, 0x6D, 0x00, 0x65, 0x00, 0x53, 0x00, 0x6C,
            0x00, 0x69, 0x00, 0x63, 0x00, 0x65, 0x00, 0x72, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00, 0x6C, 0x00, 0x65, 0x00, 0x4C, 0x00, 0x69, 0x00,
            0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0xB3, 0x10, 0x00, 0xB1, 0x10, 0x00, 0x24, 0x00, 0x97, 0x02, 0x00
        ]);

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

    _registerFormat(fmtString: string): number {
        const existing = this._fmtMap.get(fmtString);
        if (existing !== undefined) return existing.xf;
        const nid = this._nextNid++;
        const xf = this._nextXf++;
        this._fmtMap.set(fmtString, { nid, xf });
        return xf;
    }

    _buildStylesBin(): Uint8Array<ArrayBuffer> {
        if (this._fmtMap.size === 0) return this._stylesBin;
        const base = this._stylesBin;

        let fontEndOff = -1;
        for (let i = 3; i < base.length - 3; i++) {
            if (base[i] === 0xE8 && base[i + 1] === 0x04 && base[i + 2] === 0x00) { fontEndOff = i; break; }
        }
        let xfBeginOff = -1;
        for (let i = fontEndOff + 3; i < base.length - 2; i++) {
            if (base[i] === 0xE9 && base[i + 1] === 0x04) { xfBeginOff = i; break; }
        }
        let xfEndOff = -1;
        for (let i = xfBeginOff; i < base.length - 3; i++) {
            if (base[i] === 0xEA && base[i + 1] === 0x04 && base[i + 2] === 0x00) { xfEndOff = i; break; }
        }
        let fillBeginOff = -1;
        for (let i = fontEndOff + 3; i < base.length - 2; i++) {
            if (base[i] === 0xE3 && base[i + 1] === 0x04) { fillBeginOff = i; break; }
        }

        const tf = 2 + this._fmtMap.size;
        const tx = 4 + this._fmtMap.size;
        const p: Uint8Array<ArrayBuffer>[] = [];

        function vlq(v: number): number[] { const r = []; while (v >= 128) { r.push((v & 127) | 128); v >>>= 7; } r.push(v & 127); return r; }
        function brt(t: number, d: Uint8Array<ArrayBuffer>): Uint8Array<ArrayBuffer> { return new Uint8Array([...vlq(t), ...vlq(d.length), ...d]); }
        function stFmt(ifmt: number, fs: string): Uint8Array<ArrayBuffer> {
            const cch = fs.length, rem = 2 + 4 + cch * 2;
            const buf = new Uint8Array(2 + rem);
            buf[0] = 0x2C; buf[1] = rem;
            const dv = new DataView(buf.buffer);
            dv.setUint16(2, ifmt, true); dv.setUint32(4, cch, true);
            for (let i = 0; i < cch; i++) { const cp = fs.charCodeAt(i); buf[8 + i * 2] = cp & 0xFF; buf[8 + i * 2 + 1] = (cp >> 8) & 0xFF; }
            return buf;
        }
        function xfRec(fid: number, ifmt: number, fl: number, be = 0): Uint8Array<ArrayBuffer> {
            const buf = new Uint8Array(18);
            const dv = new DataView(buf.buffer);
            buf[0] = 0x2F; buf[1] = 16;
            dv.setUint16(2, fid, true); dv.setUint16(4, ifmt, true);
            buf[6] = be; buf[14] = 0x10; buf[15] = 0x10;
            dv.setUint16(16, fl, true);
            return buf;
        }

        p.push(base.subarray(0, 3));
        const fh = new Uint8Array(4); new DataView(fh.buffer).setUint16(0, tf, true);
        p.push(brt(0x0267, fh));
        p.push(stFmt(164, 'yyyy\\-mm\\-dd\\ hh:mm:ss'));
        p.push(stFmt(166, 'yyyy\\-mm\\-dd'));
        for (const [fs, st] of this._fmtMap) p.push(stFmt(st.nid, fs));
        p.push(brt(0x0268, new Uint8Array(0)));
        p.push(base.subarray(fillBeginOff, xfBeginOff));
        const xh = new Uint8Array(4); new DataView(xh.buffer).setUint16(0, tx, true);
        p.push(brt(0x0269, xh));
        p.push(xfRec(0, 0, 0x0000));
        p.push(xfRec(0, 164, 0x0001));
        p.push(xfRec(0, 166, 0x0001));
        p.push(xfRec(1, 0, 0x0000, 1));
        const sortedXf = [...this._fmtMap.entries()].sort((a, b) => a[1].xf - b[1].xf);
        for (const [, st] of sortedXf) p.push(xfRec(0, st.nid, 0x0001));
        p.push(brt(0x026A, new Uint8Array(0)));
        p.push(base.subarray(xfEndOff + 3));

        let tl = 0; for (const c of p) tl += c.length;
        const r = new Uint8Array(tl); let pos = 0;
        for (const c of p) { r.set(c, pos); pos += c.length; }
        return r;
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
        this._isStreaming = true;
        this._currentSheetBuffer = new BrowserBigBuffer();
        this._currentSheetRowNum = 0;
        this._currentSheetStartCol = 0;
        this._currentSheetEndCol = columnCount;
        this._currentSheetDoAutofilter = doAutofilter && headers !== undefined;
        this._colWidths = new Array(columnCount).fill(-1.0);
        
        if (headers) {
            for (let i = 0; i < columnCount; i++) {
                const len = headers[i] ? headers[i].length : 0;
                let w = 1.3 * len + 3;
                if (w > 80) w = 80;
                if (this._colWidths[i] < w) this._colWidths[i] = w;
            }
        }
        this._writeSheetHeader(this._currentSheetBuffer, this._currentSheetStartCol, this._currentSheetEndCol, this._currentSheetDoAutofilter);
        
        if (headers) {
            this.createRowHeader(this._currentSheetBuffer, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
            for (let c = 0; c < headers.length; c++) this.writeString(this._currentSheetBuffer, headers[c], c, hs);
            this._currentSheetRowNum++;
        }
    }

    writeRow(row: CellValue[]): void {
        const bb = this._currentSheetBuffer;
        if (!this._isStreaming || bb === null) throw new Error('Not streaming');
        if (row.length !== (this._currentSheetEndCol - this._currentSheetStartCol)) {
            throw new Error(`Row length mismatch. Expected ${this._currentSheetEndCol - this._currentSheetStartCol} columns, got ${row.length}`);
        }
        assertRowCount(this._currentSheetRowNum + 1);
        this.createRowHeader(bb, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
        for (let c = 0; c < row.length; c++) this._writeCell(bb, row[c], c);
        this._currentSheetRowNum++;
    }

    endSheet(): void {
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
        
        bb.write(this._sheet1Bytes.subarray(290));
        this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
        this._isStreaming = false;
        this._currentSheetBuffer = null;
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
        this._colWidths = new Array(colCount).fill(-1.0);
        
        if (headers) {
            for (let i = 0; i < colCount; i++) {
                const len = headers[i] ? headers[i].length : 0;
                let w = 1.3 * len + 3;
                if (w > 80) w = 80;
                if (this._colWidths[i] < w) this._colWidths[i] = w;
            }
        }
        for (let r = 0; r < Math.min(rows.length, 100); r++) {
            for (let c = 0; c < rows[r].length; c++) {
                const raw = rows[r][c]; if (raw == null) continue;
                const v = isFormattedCell(raw) ? raw.value : raw;
                if (v === null || v === undefined) continue;
                const len = (v instanceof Date) ? 10 : v.toString().length;
                let w = 1.3 * len + 3;
                if (w > 80) w = 80;
                if (this._colWidths[c] < w) this._colWidths[c] = w;
            }
        }
        
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
                startRow: 0, endRow: rn, startColumn: 0, endColumn: colCount - 1
            };
        }
        
        bb.write(this._sheet1Bytes.subarray(290));
        this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
    }

    _writeSheetHeader(bb: BrowserBigBuffer, startCol: number, endCol: number, doAf: boolean): void {
        const sh = new Uint8Array(this._sheet1Bytes);
        const shView = new DataView(sh.buffer);
        shView.setInt32(40, startCol, true);
        shView.setInt32(44, endCol, true);
        if (this._sheetCount !== 1) sh[54] = 0x9C;
        
        bb.write(sh.subarray(0, 84));
        if (doAf) bb.write(this._stickHeaderA1bytes);
        bb.write(sh.subarray(84, 159));
        
        bb.writeByte(134); bb.writeByte(3);
        for (let i = startCol; i < endCol; i++) {
            bb.writeByte(0); bb.writeByte(60); bb.writeByte(18);
            bb.writeInt32LE(i); bb.writeInt32LE(i);
            const w = this._colWidths[i] > 0 ? this._colWidths[i] : 10;
            bb.writeInt32LE(Math.round(w * 256));
            bb.writeByte(0); bb.writeByte(0); bb.writeByte(0); bb.writeByte(0);
            bb.writeByte(2);
        }
        bb.writeByte(0); bb.writeByte(135); bb.writeByte(3); bb.writeByte(0);
        bb.write(sh.subarray(159, 175));
        bb.write(new Uint8Array([38, 0]));
    }

    _writeCell(bb: BrowserBigBuffer, raw: CellValue, c: number): void {
        if (raw === null || raw === undefined) return;
        const fmtString = getFormat(raw);
        const val = unwrapCell(raw);
        const styleNum = fmtString !== null ? this._registerFormat(fmtString) : 0;

        if (typeof val === 'number') {
            if (Number.isInteger(val) && val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
                this.writeRkNumberInteger(bb, val, c, styleNum);
            } else {
                this.writeDouble(bb, val, c, styleNum);
            }
        } else if (typeof val === 'bigint') {
            this.writeString(bb, val.toString(), c);
        } else if (typeof val === 'boolean') {
            this.writeBool(bb, val, c);
        } else if (val instanceof Date) {
            const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
            this.writeDouble(bb, oaDate, c, fmtString !== null ? styleNum : 1);
        } else if (val !== null && val !== undefined) {
            this.writeString(bb, val.toString(), c, 0, fmtString !== null ? styleNum : undefined);
        }
    }

    createRowHeader(bb: BrowserBigBuffer, rn: number, sc: number, ec: number): void {
        bb.ensureCapacity(27);
        bb.writeUnsafeByte(0); bb.writeUnsafeByte(25);
        bb.writeUnsafeInt32LE(rn); bb.writeUnsafeInt32LE(0);
        bb.writeUnsafeByte(44); bb.writeUnsafeByte(1); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0);
        bb.writeUnsafeByte(0); bb.writeUnsafeByte(1); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0);
        bb.writeUnsafeInt32LE(sc); bb.writeUnsafeInt32LE(ec);
    }
    writeRkNumberInteger(bb: BrowserBigBuffer, val: number, colNum: number, styleNum = 0): void {
        bb.ensureCapacity(14);
        bb.writeUnsafeByte(2); bb.writeUnsafeByte(12);
        bb.writeUnsafeInt32LE(colNum); bb.writeUnsafeByte(styleNum); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0);
        bb.writeUnsafeInt32LE((val << 2) | 2);
    }
    writeDouble(bb: BrowserBigBuffer, val: number, colNum: number, styleNum = 0): void {
        bb.ensureCapacity(18);
        bb.writeUnsafeByte(5); bb.writeUnsafeByte(16);
        bb.writeUnsafeInt32LE(colNum); bb.writeUnsafeByte(styleNum); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0);
        bb.writeUnsafeDoubleLE(val);
    }
    writeBool(bb: BrowserBigBuffer, val: boolean, colNum: number): void {
        bb.ensureCapacity(13);
        bb.writeUnsafeByte(0x04); bb.writeUnsafeByte(9);
        bb.writeUnsafeInt32LE(colNum); bb.writeUnsafeInt32LE(0); bb.writeUnsafeByte(val ? 1 : 0);
    }
    writeDateTime(bb: BrowserBigBuffer, date: Date, colNum: number): void {
        this.writeDouble(bb, timeToOaDate(date.getTime(), this._oaEpoch), colNum, 1);
    }
    writeString(bb: BrowserBigBuffer, val: string, colNum: number, styleIndex = 0, styleOverride?: number): void {
        assertCellTextLength(val);
        let index = this._sstDic.get(val);
        if (index === undefined) {
            index = this._sstCntUnique++;
            this._sstDic.set(val, index);
        }
        this._sstCntAll++;
        const finalStyle = styleOverride !== undefined ? styleOverride : styleIndex;
        bb.ensureCapacity(17);
        bb.writeUnsafeByte(7); bb.writeUnsafeByte(12);
        bb.writeUnsafeInt32LE(colNum); bb.writeUnsafeByte(finalStyle);
        bb.writeUnsafeByte(0); bb.writeUnsafeByte(0); bb.writeUnsafeByte(0);
        bb.writeUnsafeInt32LE(index);
    }

    _saveSst(): void {
        const bb = new BrowserBigBuffer();
        bb.writeByte(159); bb.writeByte(1); bb.writeByte(8);
        bb.writeInt32LE(this._sstCntAll); bb.writeInt32LE(this._sstCntUnique);
        
        for (const [txt] of this._sstDic) {
            const tl = txt.length;
            bb.writeByte(19);
            const rl = 5 + 2 * tl;
            if (rl >= 128) {
                bb.writeByte(128 + (rl % 128));
                const tmp = rl >> 7;
                if (tmp >= 256) bb.writeByte(128 + (tmp % 128)); else bb.writeByte(tmp);
                bb.writeByte(rl >> 14);
                if ((rl >> 14) > 0) bb.writeByte(0);
            } else {
                bb.writeByte(rl & 0xFF); bb.writeByte((rl >> 8) & 0xFF);
            }
            bb.writeInt32LE(tl); bb.writeUtf16LE(txt);
        }
        bb.writeByte(160); bb.writeByte(1); bb.writeByte(0);
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
        
        let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>\n<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n<Default Extension="xml" ContentType="application/xml"/>`;
        for (const s of this._sheetList) { ct += `<Override PartName="/${s.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>\n<Override PartName="/xl/worksheets/binaryIndex${s.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>\n`; }
        ct += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>\n<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>\n</Types>`;
        this._zip.addFile('[Content_Types].xml', ct);
        
        let wr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n`;
        for (const s of this._sheetList) wr += `<Relationship Id="rId${s.sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.nameInArchive}"/>\n`;
        wr += `<Relationship Id="rId${this._sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>\n<Relationship Id="rId${this._sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>\n</Relationships>`;
        this._zip.addFile('xl/_rels/workbook.bin.rels', wr);
        
        this._zip.addFile('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>\n</Relationships>`);
        
        for (const s of this._sheetList) {
            this._zip.addFile(`xl/worksheets/_rels/${s.nameInArchive}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${s.sheetId}.bin"/>\n</Relationships>`);
        }
        
        return this._zip.toBlob();
    }
}
export default BrowserXlsbWriter;
