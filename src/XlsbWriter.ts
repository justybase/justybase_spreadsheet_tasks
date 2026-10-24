import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import archiver from 'archiver';
import { Readable } from 'stream';
import { BigBuffer } from './BigBuffer';
import { CellValue, unwrapCell, getFormat } from './Formats';
import { StreamingSheetState } from './StreamingSheetState';
import {
    sanitizeSheetName,
    initColWidths,
    applyHeaderWidths,
    updateColWidthsFromRows,
    defaultColWidth,
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
    timeToOaDate,
    assertCompressionLevel,
    DEFAULT_COMPRESSION_LEVEL,
} from './writerHelpers';

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
    private _nextFmtId: number = 167;
    private _nextXfIdx: number = 4;
    private _stFmtRecords: Map<string, { numfmtId: number; xfIndex: number }> = new Map();

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


    /**
     * @param filePath Destination `.xlsb` path on disk.
     */
    /**
     * @param filePath Destination `.xlsb` path on disk.
     * @param options.compressionLevel Deflate level 0-9 (default 3).
     */
    constructor(filePath: string, options: { compressionLevel?: number } = {}) {
        const compressionLevel = assertCompressionLevel(options.compressionLevel ?? DEFAULT_COMPRESSION_LEVEL);
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

        this._stylesBin = Buffer.from([
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
            0x00, 0xE8, 0x04, 0x00, 0xE3, 0x04, 0x04, 0x01,
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
            0x04,
            0x00, 0x00, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x00,    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0xA4, 0x00,    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0xA6, 0x00,    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x01, 0x00,
            0x2F, 0x10, 0x00, 0x00, 0x00, 0x01,    0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00, 0x00,
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

    private _registerFormat(fmtString: string): number {
        if (this._stFmtRecords.has(fmtString)) {
            return this._stFmtRecords.get(fmtString)!.xfIndex;
        }
        const numfmtId = this._nextFmtId++;
        const xfIndex = this._nextXfIdx++;
        this._stFmtRecords.set(fmtString, { numfmtId, xfIndex });
        return xfIndex;
    }

    private static _buildBrtRecord(type: number, data: Buffer): Buffer {
        const header = XlsbWriter._encodeVlq(type);
        header.push(...XlsbWriter._encodeVlq(data.length));
        return Buffer.concat([Buffer.from(header), data]);
    }

    private static _encodeVlq(value: number): number[] {
        const result: number[] = [];
        while (value >= 0x80) {
            result.push((value & 0x7F) | 0x80);
            value >>>= 7;
        }
        result.push(value & 0x7F);
        return result;
    }

    private static _buildStFmtRecord(ifmt: number, fmtString: string): Buffer {
        const cch = fmtString.length;
        const remaining = 2 + 4 + cch * 2;
        const data = Buffer.alloc(2 + remaining);
        data[0] = 0x2C;
        data[1] = remaining;
        data.writeUInt16LE(ifmt, 2);
        data.writeUInt32LE(cch, 4);
        data.write(fmtString, 8, cch * 2, 'utf16le');
        return data;
    }

    private static _buildXfRecord(fontId: number, ifmt: number, flags: number, byte6Extra: number = 0): Buffer {
        const data = Buffer.alloc(18);
        data[0] = 0x2F;
        data[1] = 16;
        data.writeUInt16LE(fontId, 2);
        data.writeUInt16LE(ifmt, 4);
        data[6] = byte6Extra;
        data[14] = 0x10;
        data[15] = 0x10;
        data.writeUInt16LE(flags, 16);
        return data;
    }

    private _buildStylesBin(): Buffer {
        if (this._stFmtRecords.size === 0) {
            return this._stylesBin;
        }
        const base = this._stylesBin;

        // Find BrtEndFonts (0xE8, 0x04, 0x00)
        let fontEndOff = -1;
        for (let i = 3; i < base.length - 3; i++) {
            if (base[i] === 0xE8 && base[i + 1] === 0x04 && base[i + 2] === 0x00) {
                fontEndOff = i;
                break;
            }
        }

        // Find BrtBeginCellXFs (0xE9, 0x04)
        let xfBeginOff = -1;
        for (let i = fontEndOff + 3; i < base.length - 2; i++) {
            if (base[i] === 0xE9 && base[i + 1] === 0x04) {
                xfBeginOff = i;
                break;
            }
        }

        // Find BrtEndCellXFs (0xEA, 0x04, 0x00)
        let xfEndOff = -1;
        for (let i = xfBeginOff; i < base.length - 3; i++) {
            if (base[i] === 0xEA && base[i + 1] === 0x04 && base[i + 2] === 0x00) {
                xfEndOff = i;
                break;
            }
        }

        // Find BrtBeginFills (0xE3, 0x04) — start after BrtEndFonts
        let fillBeginOff = -1;
        for (let i = fontEndOff + 3; i < base.length - 2; i++) {
            if (base[i] === 0xE3 && base[i + 1] === 0x04) {
                fillBeginOff = i;
                break;
            }
        }

        const totalFmt = 2 + this._stFmtRecords.size;
        const totalXf = 4 + this._stFmtRecords.size;

        const parts: Buffer[] = [];

        // 1. BrtBeginStyleSheet
        parts.push(base.subarray(0, 3));

        // 2. NEW BrtFmt: VLQ(0x0267) + VLQ(4) + count(2 LE) + pad(2)
        const fmtHeader = Buffer.alloc(4);
        fmtHeader.writeUInt16LE(totalFmt, 0);
        parts.push(XlsbWriter._buildBrtRecord(0x0267, fmtHeader));

        // 3. BrtNumFmt records: 2 base + custom (ALWAYS add base formats)
        parts.push(XlsbWriter._buildStFmtRecord(164, 'yyyy\\-mm\\-dd\\ hh:mm'));
        parts.push(XlsbWriter._buildStFmtRecord(166, 'yyyy\\-mm\\-dd'));
        for (const [fmtString, st] of this._stFmtRecords) {
            parts.push(XlsbWriter._buildStFmtRecord(st.numfmtId, fmtString));
        }

        // 4. BrtEndFonts
        parts.push(XlsbWriter._buildBrtRecord(0x0268, Buffer.alloc(0)));

        // 5. Copy fills+borders+BrtFont+cellStyleXFs from template
        parts.push(base.subarray(fillBeginOff, xfBeginOff));

        // 7. NEW BrtBeginCellXFs
        const xfHeader = Buffer.alloc(4);
        xfHeader.writeUInt16LE(totalXf, 0);
        parts.push(XlsbWriter._buildBrtRecord(0x0269, xfHeader));

        // 8. Base XF records
        parts.push(XlsbWriter._buildXfRecord(0, 0, 0x0000));
        parts.push(XlsbWriter._buildXfRecord(0, 164, 0x0001));
        parts.push(XlsbWriter._buildXfRecord(0, 166, 0x0001));
        parts.push(XlsbWriter._buildXfRecord(1, 0, 0x0000, 1));

        // 9. Custom XF records
        const sortedXf = [...this._stFmtRecords.entries()].sort((a, b) => a[1].xfIndex - b[1].xfIndex);
        for (const [, st] of sortedXf) {
            parts.push(XlsbWriter._buildXfRecord(0, st.numfmtId, 0x0001));
        }

        // 10. BrtEndCellXFs
        parts.push(XlsbWriter._buildBrtRecord(0x026A, Buffer.alloc(0)));

        // 11. Copy remaining template
        parts.push(base.subarray(xfEndOff + 3));

        return Buffer.concat(parts);
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
        options: { hidden?: boolean; doAutofilter?: boolean; sampleRows?: unknown[][] } = {}
    ): void {
        assertColumnCount(columnCount);
        if (headers !== undefined) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        const { hidden = false, doAutofilter = true, sampleRows } = options;

        // Add sheet metadata
        this.addSheet(sheetName, hidden);

        this.stream.begin(columnCount, doAutofilter && headers !== undefined, new BigBuffer());
        const bigBuf = this.stream.buffer!;

        if (headers) {
            applyHeaderWidths(this.stream.colWidths, headers, columnCount);
        }
        if (sampleRows) {
            updateColWidthsFromRows(this.stream.colWidths, sampleRows);
        }

        // Write sheet header
        const sheetHeader = Buffer.from(this._sheet1Bytes);
        sheetHeader.writeInt32LE(this.stream.startCol, 40);
        sheetHeader.writeInt32LE(this.stream.endCol, 44);

        if (this.sheetCount !== 1) {
            sheetHeader[54] = 0x9C;
        }

        bigBuf.write(sheetHeader.subarray(0, 84));

        // Write sticky header if headers are present
        if (headers) {
            bigBuf.write(this._stickHeaderA1bytes);
        }

        bigBuf.write(sheetHeader.subarray(84, 159));

        // Write column definitions
        bigBuf.writeByte(134);
        bigBuf.writeByte(3);

        for (let i = this.stream.startCol; i < this.stream.endCol; i++) {
            bigBuf.writeByte(0);
            bigBuf.writeByte(60);
            bigBuf.writeByte(18);

            bigBuf.writeInt32LE(i);
            bigBuf.writeInt32LE(i);

            const width = defaultColWidth(this.stream.colWidths[i], true);
            bigBuf.writeByte(0);
            bigBuf.writeByte(Math.max(0, Math.min(255, width)));
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);

            bigBuf.writeByte(0);
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);

            bigBuf.writeByte(2);
        }

        bigBuf.writeByte(0);
        bigBuf.writeByte(135);
        bigBuf.writeByte(3);
        bigBuf.writeByte(0);

        bigBuf.write(sheetHeader.subarray(159, 175));
        bigBuf.write(Buffer.from([38, 0]));

        // Write header row if provided
        if (headers) {
            this.createRowHeader(bigBuf, this.stream.rowNum, this.stream.startCol, this.stream.endCol);
            for (let c = 0; c < headers.length; c++) {
                this.writeString(bigBuf, headers[c], c, true);
            }
            this.stream.rowNum++;
        }
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
        assertRowCount(this.stream.rowNum + 1);

        // Write row header
        this.createRowHeader(bigBuf, this.stream.rowNum, this.stream.startCol, this.stream.endCol);

        // Write each cell
        for (let c = 0; c < row.length; c++) {
            const raw = row[c];
            if (raw === null || raw === undefined) continue;

            const fmtString = getFormat(raw);
            const val = fmtString !== null ? unwrapCell(raw) : raw;
            const styleNum = fmtString !== null ? this._registerFormat(fmtString) : 0;

            if (typeof val === 'number') {
                if (Number.isInteger(val)) {
                    if (val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
                        this.writeRkNumberInteger(bigBuf, val, c, styleNum);
                    } else {
                        this.writeDouble(bigBuf, val, c, styleNum);
                    }
                } else {
                    this.writeDouble(bigBuf, val, c, styleNum);
                }
            } else if (typeof val === 'bigint') {
                this.writeString(bigBuf, val.toString(), c);
            } else if (typeof val === 'boolean') {
                this.writeBool(bigBuf, val, c);
            } else if (val instanceof Date) {
                const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
                this.writeDouble(bigBuf, oaDate, c, fmtString !== null ? styleNum : 1);
            } else if (val !== null && val !== undefined) {
                this.writeString(bigBuf, val.toString(), c, false, fmtString !== null ? styleNum : undefined);
            }
        }

        this.stream.rowNum++;
    }

    /**
     * Finalize the current streaming sheet and add it to the archive.
     * Must be called after startSheet() and writeRow() calls.
     */
    endSheet(): void {
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

        bigBuf.write(sheetHeader.subarray(290));

        // Add to archive
        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: this.sheetList[this.sheetCount - 1].pathInArchive
        });

        // Reset streaming state
        this.stream.end();
    }


    /**
     * Write an entire sheet in one call (batch mode).
     * @param rows Data rows (`CellValue` cells; null/undefined = blank)
     * @param headers Optional header row (bold + frozen + autofilter when enabled)
     * @param doAutofilter Enable autofilter when headers are present (default true)
     */
    writeSheet(rows: CellValue[][], headers: string[] | null = null, doAutofilter: boolean = true): void {
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

        const colWidths = initColWidths(columnCount);

        if (headers) {
            applyHeaderWidths(colWidths, headers, columnCount);
        }
        updateColWidthsFromRows(colWidths, rows);

        const sheetHeader = Buffer.from(this._sheet1Bytes);
        const startCol = 0;
        const endCol = columnCount;

        sheetHeader.writeInt32LE(startCol, 40);
        sheetHeader.writeInt32LE(endCol, 44);

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

        for (let i = startCol; i < endCol; i++) {
            bigBuf.writeByte(0);
            bigBuf.writeByte(60);
            bigBuf.writeByte(18);

            bigBuf.writeInt32LE(i);
            bigBuf.writeInt32LE(i);

            const width = defaultColWidth(colWidths[i], true);
            bigBuf.writeByte(0);
            bigBuf.writeByte(Math.max(0, Math.min(255, width)));
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);

            bigBuf.writeByte(0);
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);
            bigBuf.writeByte(0);

            bigBuf.writeByte(2);
        }

        bigBuf.writeByte(0);
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
                const raw = row[c];
                if (raw === null || raw === undefined) continue;

                const fmtString = getFormat(raw);
                const val = fmtString !== null ? unwrapCell(raw) : raw;
                const styleNum = fmtString !== null ? this._registerFormat(fmtString) : 0;

                if (typeof val === 'number') {
                    if (Number.isInteger(val)) {
                        if (val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
                            this.writeRkNumberInteger(bigBuf, val, c, styleNum);
                        } else {
                            this.writeDouble(bigBuf, val, c, styleNum);
                        }
                    } else {
                        this.writeDouble(bigBuf, val, c, styleNum);
                    }
                } else if (typeof val === 'bigint') {
                    this.writeString(bigBuf, val.toString(), c);
                } else if (typeof val === 'boolean') {
                    this.writeBool(bigBuf, val, c);
                } else if (val instanceof Date) {
                    const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
                    this.writeDouble(bigBuf, oaDate, c, fmtString !== null ? styleNum : 1);
                } else if (val !== null && val !== undefined) {
                    this.writeString(bigBuf, val.toString(), c, false, fmtString !== null ? styleNum : undefined);
                }
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

        bigBuf.write(sheetHeader.subarray(290));

        this.archive.append(Readable.from(bigBuf.getChunks()), { name: this.sheetList[this.sheetCount - 1].pathInArchive });
    }



    createRowHeader(bigBuf: BigBuffer, rowNumber: number, startCol: number, endCol: number): void {
        bigBuf.ensureCapacity(27);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(25);
        bigBuf.writeUnsafeInt32LE(rowNumber);
        bigBuf.writeUnsafeInt32LE(0);
        bigBuf.writeUnsafeByte(44);
        bigBuf.writeUnsafeByte(1);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(1);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeInt32LE(startCol);
        bigBuf.writeUnsafeInt32LE(endCol);
    }

    writeRkNumberInteger(bigBuf: BigBuffer, val: number, colNum: number, styleNum: number = 0): void {
        bigBuf.ensureCapacity(14);
        bigBuf.writeUnsafeByte(2);
        bigBuf.writeUnsafeByte(12);
        bigBuf.writeUnsafeInt32LE(colNum);
        bigBuf.writeUnsafeByte(styleNum);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        const rkVal = (val << 2) | 2;
        bigBuf.writeUnsafeInt32LE(rkVal);
    }

    writeDouble(bigBuf: BigBuffer, val: number, colNum: number, styleNum: number = 0): void {
        bigBuf.ensureCapacity(18);
        bigBuf.writeUnsafeByte(5);
        bigBuf.writeUnsafeByte(16);
        bigBuf.writeUnsafeInt32LE(colNum);
        bigBuf.writeUnsafeByte(styleNum);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeDoubleLE(val);
    }

    writeBool(bigBuf: BigBuffer, val: boolean, colNum: number): void {
        bigBuf.ensureCapacity(13);
        bigBuf.writeUnsafeByte(0x04);
        bigBuf.writeUnsafeByte(9);
        bigBuf.writeUnsafeInt32LE(colNum);
        bigBuf.writeUnsafeInt32LE(0);
        bigBuf.writeUnsafeByte(val ? 1 : 0);
    }

    writeDateTime(bigBuf: BigBuffer, date: Date, colNum: number): void {
        // Use UTC time to avoid timezone conversion issues; handles the
        // fictitious 1900-02-29.
        const oaDate = timeToOaDate(date.getTime(), this._oaEpoch);
        this.writeDouble(bigBuf, oaDate, colNum, 1);
    }

    private _getDateWidth(_date: Date): number {
        return 19; // "YYYY-MM-DD HH:MM:SS" ≈ 19 chars
    }

    writeString(bigBuf: BigBuffer, val: string, colNum: number, bolded: boolean = false, styleOverride?: number): void {
        assertCellTextLength(val);
        let index: number;
        if (this.sstDic.has(val)) {
            index = this.sstDic.get(val)!;
        } else {
            index = this.sstCntUnique;
            this.sstDic.set(val, index);
            this.sstCntUnique++;
        }
        this.sstCntAll++;

        const styleNum = styleOverride !== undefined ? styleOverride : (bolded ? 3 : 0);
        bigBuf.ensureCapacity(17);
        bigBuf.writeUnsafeByte(7);
        bigBuf.writeUnsafeByte(12);
        bigBuf.writeUnsafeInt32LE(colNum);
        bigBuf.writeUnsafeByte(styleNum);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeByte(0);
        bigBuf.writeUnsafeInt32LE(index);
    }

    private saveSst(): void {
        const bigBuf = new BigBuffer();

        bigBuf.writeByte(159);
        bigBuf.writeByte(1);
        bigBuf.writeByte(8);

        bigBuf.writeInt32LE(this.sstCntAll);
        bigBuf.writeInt32LE(this.sstCntUnique);

        for (const [txt] of this.sstDic) {
            const txtLen = txt.length;

            bigBuf.writeByte(19);

            const recLen = 5 + 2 * txtLen;

            if (recLen >= 128) {
                bigBuf.writeByte(128 + (recLen % 128));
                const tmp = recLen >> 7;
                if (tmp >= 256) {
                    bigBuf.writeByte(128 + (tmp % 128));
                } else {
                    bigBuf.writeByte(tmp);
                }
                bigBuf.writeByte(recLen >> 14);
                if ((recLen >> 14) > 0) {
                    bigBuf.writeByte(0);
                }
            } else {
                bigBuf.writeByte(recLen & 0xFF);
                bigBuf.writeByte((recLen >> 8) & 0xFF);
            }

            bigBuf.writeInt32LE(txtLen);
            bigBuf.writeUtf16LE(txt);
        }

        bigBuf.writeByte(160);
        bigBuf.writeByte(1);
        bigBuf.writeByte(0);

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

                this.archive.append(this._buildStylesBin(), { name: 'xl/styles.bin' });

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
                    const wsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${sheet.sheetId}.bin"/>
</Relationships>`;
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
