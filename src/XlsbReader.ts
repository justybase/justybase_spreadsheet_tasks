import AdmZip from 'adm-zip';
import { ExcelReaderAbstract } from './ExcelReaderAbstract';
import { BiffReaderWriter } from './BiffReaderWriter';
import { CellValue } from './Formats';
import { parseTagAttributes, resolveRelTarget } from './xmlUtils';
import { readRecord, Biff12Record } from './biff12Utils';
import { MAX_ROWS } from './writerHelpers';

interface SheetInfo {
    name: string;
    rId: string;
    path: string | null;
}

export class XlsbReader extends ExcelReaderAbstract {
    private zip: AdmZip | null = null;
    private sharedStrings: string[] = [];
    private sheetNames: string[] = [];
    private sheets: SheetInfo[] = [];

    private _currentSheetIndex: number = -1;
    private _reader: BiffReaderWriter | null = null;
    private _currentRow: CellValue[] = [];
    private _pendingRowIndex: number = -1;
    private _nextExpectedRow: number = 0;
    private _eof: boolean = false;

    xfIdToNumFmtId: number[] = [];
    customDateFormats: Set<number> | null = null;

    constructor() {
        super();
    }

    async open(path: string, readSharedStrings: boolean = true): Promise<void> {
        await this.close();
        // Reset state for re-open (xlspy parity).
        this.sheetNames = [];
        this.sheets = [];
        this.sharedStrings = [];
        this.xfIdToNumFmtId = [];
        this.customDateFormats = new Set();
        this.fieldCount = 0;
        this.resultsCount = 0;
        this.actualSheetName = '';
        this._currentSheetIndex = -1;
        this._currentRow = [];
        this._pendingRowIndex = -1;
        this._nextExpectedRow = 0;
        this._eof = false;
        this._reader = null;
        this._oaEpoch = Date.UTC(1899, 11, 30);
        this.zip = new AdmZip(path);

        const wbEntry = this.zip.getEntry('xl/workbook.bin');
        if (wbEntry) {
            const buf = wbEntry.getData();
            const reader = new BiffReaderWriter(buf);
            while (reader.readWorkbook()) {
                if (reader._isSheet) {
                    const name = reader._workbookName!;
                    const rId = reader._recId!;
                    this.sheetNames.push(name);
                    this.sheets.push({ name, rId, path: null });
                }
            }
        }

        const relsEntry = this.zip.getEntry('xl/_rels/workbook.bin.rels');
        if (relsEntry) {
            const xml = relsEntry.getData().toString('utf8');
            const relTagRe = /<Relationship\b[^>]*>/g;
            let tagMatch;
            const rIdToTarget: Record<string, string> = {};
            while ((tagMatch = relTagRe.exec(xml)) !== null) {
                const attrs = parseTagAttributes(tagMatch[0]);
                const id = attrs['Id'] ?? attrs['id'];
                const target = attrs['Target'] ?? attrs['target'];
                if (id !== undefined && target !== undefined) {
                    rIdToTarget[id] = target;
                }
            }

            for (const sheet of this.sheets) {
                const target = rIdToTarget[sheet.rId];
                if (target) {
                    sheet.path = resolveRelTarget(target);
                }
            }
        }

        // 1904 date system: BrtWbProp (0x99) bit0 f1904 in workbook.bin.
        try {
            const raw = wbEntry ? wbEntry.getData() : null;
            if (raw) {
                const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
                let pos = 0;
                while (readRecord(raw, pos, rec)) {
                    if (rec.id === 0x99 && rec.len >= 4) {
                        const is1904 = (raw[rec.dataStart] & 0x01) !== 0;
                        this._oaEpoch = is1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
                        break;
                    }
                    pos = rec.dataEnd;
                }
            }
        } catch {
            // Keep default epoch on parse errors.
        }

        if (readSharedStrings) {
            const ssPath = 'xl/sharedStrings.bin';
            if (this.zip.getEntry(ssPath)) {
                this._readSharedStrings(ssPath);
            }
        }

        const stylesPath = 'xl/styles.bin';
        if (this.zip.getEntry(stylesPath)) {
            this._readStyles(stylesPath);
        }

        this.resultsCount = this.sheets.length;
        this._currentSheetIndex = -1;
    }

    private _readSharedStrings(path: string): void {
        const entry = this.zip!.getEntry(path);
        if (!entry) return;
        const reader = new BiffReaderWriter(entry.getData());

        while (reader.readSharedStrings()) {
            if (reader._sharedStringValue !== null) {
                this.sharedStrings.push(reader._sharedStringValue);
            }
        }
    }

    private _readStyles(path: string): void {
        const entry = this.zip!.getEntry(path);
        if (!entry) return;

        const reader = new BiffReaderWriter(entry.getData());
        while (reader.readStyles()) {
            // just consume
        }

        this.xfIdToNumFmtId = reader._xfIndexToNumFmtId;
        this.customDateFormats = reader._customNumFmts;
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

        if (this._eof) return false;

        // Gap rows (missing BrtRowHdr) are returned as empty rows. An index beyond
        // Excel's row limit is malformed: skip the fill instead of yielding millions of rows.
        if (this._pendingRowIndex > this._nextExpectedRow && this._pendingRowIndex < MAX_ROWS) {
            this._currentRow = [];
            this._nextExpectedRow++;
            return true;
        }

        this._currentRow = [];

        while (true) {
            const hasRecord = this._reader!.readWorksheet();
            if (!hasRecord) {
                this._eof = true;
                // If the current row never had cells and we never yielded it,
                // still return it once (handles single empty row sheets).
                // When _currentRow is empty and pending was already yielded,
                // the caller already got it via gap logic; avoid duplicate.
                // Here pending == expected, so this is the last row.
                this._nextExpectedRow = this._pendingRowIndex + 1;
                return true;
            }

            if (this._reader!._rowIndex !== -1 && this._reader!._rowIndex !== this._pendingRowIndex) {
                const completed = this._pendingRowIndex;
                this._pendingRowIndex = this._reader!._rowIndex;
                this._nextExpectedRow = completed + 1;
                return true;
            }

            if (this._reader!._readCell) {
                const col = this._reader!._columnNum;
                let val: CellValue = null;
                switch (this._reader!._cellType) {
                    case 2:
                        val = this.sharedStrings[this._reader!._intValue];
                        break;
                    case 3: {
                        val = this._reader!._doubleVal;

                        const xfIndex = this._reader!._xfIndex;
                        let numFmtId = 0;
                        if (this.xfIdToNumFmtId && xfIndex < this.xfIdToNumFmtId.length) {
                            numFmtId = this.xfIdToNumFmtId[xfIndex];
                        }

                        const isDate = (numFmtId >= 14 && numFmtId <= 22) ||
                            (numFmtId >= 45 && numFmtId <= 47) ||
                            (this.customDateFormats && this.customDateFormats.has(numFmtId));

                        if (isDate) {
                            try {
                                val = this.getDateTimeFromOaDate(val as number);
                            } catch {
                                // ignore date conversion errors
                            }
                        }
                        break;
                    }
                    case 4:
                        val = this._reader!._boolValue;
                        break;
                    case 5:
                        val = this._reader!._stringValue;
                        break;
                    default:
                        val = null;
                }

                this._currentRow[col] = val;
                if (col >= this.fieldCount) this.fieldCount = col + 1;
            }
        }
    }

    private _initSheet(index: number): boolean {
        if (index < 0 || index >= this.sheets.length) return false;
        const sheet = this.sheets[index];
        const entry = sheet.path ? this.zip!.getEntry(sheet.path) : undefined;
        // Only a missing part is an error; a valid but empty sheet is selected
        // successfully and reports EOF through read().
        if (!entry) return false;

        this._reader = new BiffReaderWriter(entry.getData());
        this._eof = false;
        this._pendingRowIndex = -1;
        this._nextExpectedRow = 0;
        this._currentRow = [];
        this.fieldCount = 0;
        this.actualSheetName = sheet.name;

        while (this._reader.readWorksheet()) {
            if (this._reader._rowIndex !== -1) {
                this._pendingRowIndex = this._reader._rowIndex;
                return true;
            }
        }

        this._eof = true;
        return true;
    }

    getValue(i: number): CellValue {
        if (i < 0 || i >= this._currentRow.length) return null;
        return this._currentRow[i];
    }
}

export default XlsbReader;
