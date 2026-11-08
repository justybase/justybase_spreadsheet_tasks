import AdmZip from 'adm-zip';
import { ExcelReaderAbstract } from './ExcelReaderAbstract';
import { BiffReaderWriter } from './BiffReaderWriter';
import { CellStyle, CellValue } from './Formats';
import { parseXlsbStylesBin } from './writers/core/xlsbStyles';
import { parseTagAttributes, resolveRelTarget } from './xmlUtils';
import { biffCodeToExcelError, parseSharedStringsBin, readRecord, Biff12Record } from './biff12Utils';
import { MAX_COLS, MAX_ROWS, oaDateToTime } from './writerHelpers';

interface SheetInfo {
    name: string;
    rId: string;
    path: string | null;
}

/** Cell record ids in a worksheet part. */
const BRT_CELL_BLANK = 0x01;
const BRT_CELL_RK = 0x02;
const BRT_CELL_ERROR = 0x03;
const BRT_CELL_BOOL = 0x04;
const BRT_CELL_REAL = 0x05;
const BRT_CELL_ST = 0x06;
const BRT_CELL_ISST = 0x07;
const BRT_FMLA_STRING = 0x08;
const BRT_FMLA_NUM = 0x09;
const BRT_FMLA_BOOL = 0x0a;
const BRT_FMLA_ERROR = 0x0b;

export class XlsbReader extends ExcelReaderAbstract {
    private zip: AdmZip | null = null;
    private sharedStrings: string[] = [];
    private sheetNames: string[] = [];
    private sheets: SheetInfo[] = [];

    private _currentSheetIndex: number = -1;
    private _sheetBuffer: Buffer | null = null;
    private _pos: number = 0;
    private _currentRow: CellValue[] = [];
    private _currentStyles: number[] = [];
    private _pendingRowIndex: number = -1;
    private _nextExpectedRow: number = 0;
    private _eof: boolean = false;
    private readonly _rkScratch: Buffer = Buffer.allocUnsafe(8);

    xfIdToNumFmtId: number[] = [];
    customDateFormats: Set<number> | null = null;
    /** One flag per xf id: 1 when its number format is a date. */
    private _dateXf: Uint8Array = new Uint8Array(0);
    private _xfStyles: Array<CellStyle | null> = [];

    constructor() {
        super();
    }

    async open(path: string, readSharedStrings: boolean = true): Promise<void> {
        await this.close();
        this._resetState();
        this.zip = new AdmZip(path);
        this._load(readSharedStrings);
    }

    /** Open a workbook already held in memory (ZIP bytes). */
    async openBuffer(buffer: Buffer, readSharedStrings: boolean = true): Promise<void> {
        await this.close();
        this._resetState();
        this.zip = new AdmZip(buffer);
        this._load(readSharedStrings);
    }

    private _resetState(): void {
        this.sheetNames = [];
        this.sheets = [];
        this.sharedStrings = [];
        this.xfIdToNumFmtId = [];
        this.customDateFormats = new Set();
        this._dateXf = new Uint8Array(0);
        this._xfStyles = [];
        this.fieldCount = 0;
        this.resultsCount = 0;
        this.actualSheetName = '';
        this._currentSheetIndex = -1;
        this._sheetBuffer = null;
        this._pos = 0;
        this._currentRow = [];
        this._currentStyles = [];
        this._pendingRowIndex = -1;
        this._nextExpectedRow = 0;
        this._eof = false;
        this._oaEpoch = Date.UTC(1899, 11, 30);
    }

    private _load(readSharedStrings: boolean): void {
        const wbEntry = this.zip!.getEntry('xl/workbook.bin');
        let workbookBuffer: Buffer | null = null;
        if (wbEntry) {
            workbookBuffer = wbEntry.getData();
            const reader = new BiffReaderWriter(workbookBuffer);
            while (reader.readWorkbook()) {
                if (reader._isSheet) {
                    const name = reader._workbookName!;
                    const rId = reader._recId!;
                    this.sheetNames.push(name);
                    this.sheets.push({ name, rId, path: null });
                }
            }
        }

        const relsEntry = this.zip!.getEntry('xl/_rels/workbook.bin.rels');
        if (relsEntry) {
            const xml = relsEntry.getData().toString('utf8');
            const relTagRe = /<Relationship\b[^>]*>/g;
            let tagMatch;
            const rIdToTarget: Record<string, string> = {};
            while ((tagMatch = relTagRe.exec(xml)) !== null) {
                const attrs = parseTagAttributes(tagMatch[0]);
                const id = attrs['Id'] ?? attrs['id'];
                const target = attrs['Target'] ?? attrs['target'];
                if (id !== undefined && target !== undefined) rIdToTarget[id] = target;
            }
            for (const sheet of this.sheets) {
                const target = rIdToTarget[sheet.rId];
                if (target) sheet.path = resolveRelTarget(target);
            }
        }

        // 1904 date system: BrtWbProp (0x99) bit0 f1904 in workbook.bin.
        if (workbookBuffer) {
            try {
                const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
                let pos = 0;
                while (readRecord(workbookBuffer, pos, rec)) {
                    if (rec.id === 0x99 && rec.len >= 4) {
                        const is1904 = (workbookBuffer[rec.dataStart] & 0x01) !== 0;
                        this._oaEpoch = is1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
                        break;
                    }
                    pos = rec.dataEnd;
                }
            } catch {
                // Keep the default epoch on parse errors.
            }
        }

        if (readSharedStrings) {
            const ssEntry = this.zip!.getEntry('xl/sharedStrings.bin');
            if (ssEntry) {
                this.sharedStrings = parseSharedStringsBin(ssEntry.getData()).values;
            }
        }

        const stylesEntry = this.zip!.getEntry('xl/styles.bin');
        if (stylesEntry) {
            const reader = new BiffReaderWriter(stylesEntry.getData());
            while (reader.readStyles()) {
                // consume
            }
            this.xfIdToNumFmtId = reader._xfIndexToNumFmtId;
            this.customDateFormats = reader._customNumFmts;
            this._xfStyles = parseXlsbStylesBin(stylesEntry.getData());
        }
        this._buildDateXf();

        this.resultsCount = this.sheets.length;
        this._currentSheetIndex = -1;
    }

    private _buildDateXf(): void {
        this._dateXf = new Uint8Array(this.xfIdToNumFmtId.length);
        for (let i = 0; i < this.xfIdToNumFmtId.length; i++) {
            const nf = this.xfIdToNumFmtId[i];
            const isDate = (nf >= 14 && nf <= 22) ||
                (nf >= 45 && nf <= 47) ||
                (this.customDateFormats !== null && this.customDateFormats.has(nf));
            this._dateXf[i] = isDate ? 1 : 0;
        }
    }

    private _isDateStyle(xf: number): boolean {
        return xf > 0 && xf < this._dateXf.length && this._dateXf[xf] !== 0;
    }

    getSheetNames(): string[] {
        return this.sheetNames;
    }

    async selectSheet(nameOrIndex: string | number): Promise<void> {
        let index = -1;
        if (typeof nameOrIndex === 'number') {
            index = nameOrIndex;
        } else {
            index = this.sheetNames.indexOf(nameOrIndex);
            if (index === -1) throw new Error(`Sheet '${nameOrIndex}' not found`);
        }
        if (index < 0 || index >= this.sheets.length) {
            throw new Error(`Sheet index ${String(nameOrIndex)} out of range`);
        }
        this._currentSheetIndex = index;
        if (!this._initSheet(index)) {
            throw new Error(`XlsbReader: worksheet part missing for sheet '${this.sheetNames[index]}'.`);
        }
    }

    async nextResult(): Promise<boolean> {
        const next = this._currentSheetIndex + 1;
        if (next >= this.sheets.length) return false;
        this._currentSheetIndex = next;
        if (!this._initSheet(next)) {
            throw new Error(`XlsbReader: worksheet part missing for sheet '${this.sheetNames[next]}'.`);
        }
        return true;
    }

    async read(): Promise<boolean> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            if (!this._initSheet(0)) return false;
        }
        return this._readNext();
    }

    /**
     * Synchronous row access for batch consumers. Returns the internal row
     * array, valid until the next call; `readAll`/`rows` return copies.
     * Select a sheet (or call `read()` once) before using it.
     */
    readRow(): CellValue[] | null {
        if (this._currentSheetIndex === -1) {
            throw new Error('XlsbReader: select a sheet before readRow()');
        }
        return this._readNext() ? this._currentRow : null;
    }

    /** Collect every remaining row of the current sheet. */
    async readAll(): Promise<CellValue[][]> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            if (!this._initSheet(0)) return [];
        }
        const rows: CellValue[][] = [];
        while (true) {
            const row = this.readRow();
            if (row === null) break;
            rows.push(row.slice());
        }
        return rows;
    }

    /** Stream every remaining row of the current sheet. */
    async *rows(): AsyncGenerator<CellValue[]> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            if (!this._initSheet(0)) return;
        }
        while (true) {
            const row = this.readRow();
            if (row === null) return;
            yield row.slice();
        }
    }

    private _readNext(): boolean {
        if (this._eof) return false;
        const buffer = this._sheetBuffer;
        if (buffer === null) return false;

        // Gap rows (missing BrtRowHdr) are returned as empty rows. An index
        // beyond Excel's row limit is malformed: do not expand it.
        if (this._pendingRowIndex > this._nextExpectedRow && this._pendingRowIndex < MAX_ROWS) {
            this._currentRow = [];
            this._currentStyles = [];
            this._nextExpectedRow++;
            return true;
        }

        const row: CellValue[] = [];
        const styles: number[] = [];
        this._currentStyles = styles;
        let fieldCount = this.fieldCount;

        while (this._pos < buffer.length) {
            let id = 0;
            let shift = 0;
            let byte = 0;
            do {
                if (this._pos >= buffer.length) return this._finishEof(row, fieldCount);
                byte = buffer[this._pos++];
                id |= (byte & 0x7f) << shift;
                shift += 7;
            } while ((byte & 0x80) !== 0);

            let len = 0;
            shift = 0;
            do {
                if (this._pos >= buffer.length) return this._finishEof(row, fieldCount);
                byte = buffer[this._pos++];
                len |= (byte & 0x7f) << shift;
                shift += 7;
            } while ((byte & 0x80) !== 0);

            const body = this._pos;
            if (body + len > buffer.length) return this._finishEof(row, fieldCount);
            this._pos = body + len;

            if (id === 0x00) {
                // BrtRowHdr
                const rowIndex = len >= 4 ? buffer.readInt32LE(body) : -1;
                const normalized = rowIndex < 0 ? -1 : rowIndex;
                if (normalized !== -1 && normalized !== this._pendingRowIndex) {
                    const completed = this._pendingRowIndex;
                    this._pendingRowIndex = normalized;
                    this._nextExpectedRow = completed + 1;
                    this._currentRow = row;
                    this._currentStyles = styles;
                    this.fieldCount = fieldCount;
                    return true;
                }
                continue;
            }

            if (id < BRT_CELL_BLANK || id > BRT_FMLA_ERROR || len < 8) continue;
            const col = buffer.readUInt32LE(body);
            if (col >= MAX_COLS) continue;
            const xf = buffer.readUInt32LE(body + 4) & 0xffffff;
            let value: CellValue = null;

            switch (id) {
                case BRT_CELL_BLANK:
                    break;
                case BRT_CELL_RK:
                    if (len >= 12) value = this._numberValue(buffer.readInt32LE(body + 8), xf);
                    break;
                case BRT_CELL_ERROR:
                case BRT_FMLA_ERROR:
                    if (len >= 9) value = biffCodeToExcelError(buffer[body + 8]);
                    break;
                case BRT_CELL_BOOL:
                case BRT_FMLA_BOOL:
                    if (len >= 9) value = buffer[body + 8] === 1;
                    break;
                case BRT_CELL_REAL:
                case BRT_FMLA_NUM:
                    if (len >= 16) value = this._doubleValue(buffer.readDoubleLE(body + 8), xf);
                    break;
                case BRT_CELL_ST:
                case BRT_FMLA_STRING: {
                    if (len >= 12) {
                        const chars = buffer.readUInt32LE(body + 8);
                        if (chars * 2 + 12 <= len) {
                            value = buffer.toString('utf16le', body + 12, body + 12 + chars * 2);
                        }
                    }
                    break;
                }
                case BRT_CELL_ISST: {
                    if (len >= 12) {
                        const index = buffer.readUInt32LE(body + 8);
                        value = this.sharedStrings[index] ?? null;
                    }
                    break;
                }
            }

            row[col] = value;
            styles[col] = xf;
            if (col + 1 > fieldCount) fieldCount = col + 1;
        }

        return this._finishEof(row, fieldCount);
    }

    private _finishEof(row: CellValue[] = this._currentRow, fieldCount: number = this.fieldCount): boolean {
        this._eof = true;
        this._nextExpectedRow = this._pendingRowIndex + 1;
        this._currentRow = row;
        this.fieldCount = fieldCount;
        return true;
    }

    private _numberValue(raw: number, xf: number): CellValue {
        let value: number;
        if ((raw & 0x02) !== 0) {
            value = raw >> 2;
        } else {
            this._rkScratch.writeInt32LE(0, 0);
            this._rkScratch.writeInt32LE(raw & 0xfffffffc, 4);
            value = this._rkScratch.readDoubleLE(0);
        }
        if ((raw & 0x01) !== 0) value /= 100;
        return this._isDateStyle(xf) ? this._asDate(value) : value;
    }

    private _doubleValue(value: number, xf: number): CellValue {
        return this._isDateStyle(xf) ? this._asDate(value) : value;
    }

    private _asDate(value: number): CellValue {
        try {
            return new Date(oaDateToTime(value, this._oaEpoch));
        } catch {
            return value;
        }
    }

    private _initSheet(index: number): boolean {
        if (index < 0 || index >= this.sheets.length) return false;
        const sheet = this.sheets[index];
        const entry = sheet.path ? this.zip!.getEntry(sheet.path) : undefined;
        if (!entry) return false;

        this._sheetBuffer = entry.getData();
        this._pos = 0;
        this._eof = false;
        this._pendingRowIndex = -1;
        this._nextExpectedRow = 0;
        this._currentRow = [];
        this._currentStyles = [];
        this.fieldCount = 0;
        this.actualSheetName = sheet.name;

        const buffer = this._sheetBuffer;
        while (this._pos < buffer.length) {
            let id = 0;
            let shift = 0;
            let byte = 0;
            do {
                if (this._pos >= buffer.length) break;
                byte = buffer[this._pos++];
                id |= (byte & 0x7f) << shift;
                shift += 7;
            } while ((byte & 0x80) !== 0);
            let len = 0;
            shift = 0;
            do {
                if (this._pos >= buffer.length) break;
                byte = buffer[this._pos++];
                len |= (byte & 0x7f) << shift;
                shift += 7;
            } while ((byte & 0x80) !== 0);
            const body = this._pos;
            if (body + len > buffer.length) break;
            this._pos = body + len;
            if (id === 0x00 && len >= 4) {
                this._pendingRowIndex = buffer.readInt32LE(body);
                return true;
            }
        }

        this._eof = true;
        return true;
    }

    /** Style of the current row cell at column `i`, or null for the default style. */
    getStyle(i: number): CellStyle | null {
        const xf = this._currentStyles[i];
        if (xf === undefined) return null;
        return this._xfStyles[xf] ?? null;
    }

    getValue(i: number): CellValue {
        if (i < 0 || i >= this._currentRow.length) return null;
        return this._currentRow[i];
    }
}

export default XlsbReader;
