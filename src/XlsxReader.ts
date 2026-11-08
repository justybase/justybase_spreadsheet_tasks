import yauzl, { ZipFile, Entry } from 'yauzl';
import { ExcelReaderAbstract } from './ExcelReaderAbstract';
import { CellStyle, CellValue, RichText } from './Formats';
import {
    parseXlsxInlineRichText,
    parseXlsxRichTexts,
    parseXlsxStyles,
} from './xlsxStyleParser';
import {
    unescapeXml,
    unescapeXmlText,
    columnLetterToIndex,
    parseSharedStringsBuffer,
    parseSharedStringsXml,
    parseTagAttributes,
    resolveRelTarget,
} from './xmlUtils';
import { MAX_ROWS } from './writerHelpers';

interface SheetInfo {
    name: string;
    sheetId: string;
    rId: string;
    path: string;
}

/**
 * Parse an ISO 8601 `t="d"` cell as UTC. Numeric OA dates are converted from a
 * UTC epoch, so a value without an offset must not be interpreted as local time.
 */
function parseIsoDateUtc(value: string): Date | null {
    let text = value.trim().replace(' ', 'T');
    if (!text) return null;
    if (!/z$|[+-]\d{2}:?\d{2}$/i.test(text) && text.includes('T')) {
        text += 'Z';
    }
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
}

/** Index of the `>` that closes a tag starting at `start`, or -1. */
function tagEndIndex(buf: Buffer, start: number, limit: number): number {
    const end = buf.indexOf(0x3e, start);
    return end === -1 || end >= limit ? -1 : end;
}

const BYTE_LT = 0x3c;
const BYTE_GT = 0x3e;
const BYTE_SLASH = 0x2f;

/** Index of byte `code` in `[start, end)`, or -1. */
function indexOfByte(buf: Buffer, code: number, start: number, end: number): number {
    for (let i = start; i < end; i++) {
        if (buf[i] === code) return i;
    }
    return -1;
}

/**
 * Index of `<name` in `[start, end)`, or -1.
 *
 * A byte loop beats `Buffer#subarray(...).indexOf` for the few-byte cell
 * bodies: it allocates nothing and stays inside the cell.
 */
function indexOfTag(buf: Buffer, name: number, start: number, end: number): number {
    for (let i = start; i + 1 < end; i++) {
        if (buf[i] === BYTE_LT && buf[i + 1] === name) return i;
    }
    return -1;
}

/** Index of `</name>` in `[start, end)`, or -1. */
function indexOfCloseTag(buf: Buffer, name: number, start: number, end: number): number {
    for (let i = start; i + 3 < end; i++) {
        if (buf[i] === BYTE_LT && buf[i + 1] === BYTE_SLASH &&
            buf[i + 2] === name && buf[i + 3] === BYTE_GT) {
            return i;
        }
    }
    return -1;
}

/** String-path counterpart of {@link indexOfCloseTag}. */
function stringIndexOfWithin(xml: string, needle: string, start: number, end: number): number {
    if (end <= start) return -1;
    const index = xml.substring(start, end).indexOf(needle);
    return index === -1 ? -1 : start + index;
}

/** Value of attribute `name` in `<tag ...>`; raw (entities are not decoded). */
function attributeValue(buf: Buffer, start: number, end: number, name: string): string | null {
    let pos = start + 1;
    // Skip the element name: `<c t="s">` starts with the tag, not an attribute.
    while (pos < end && buf[pos] > 0x20 && buf[pos] !== 0x3e && buf[pos] !== 0x2f) pos++;
    while (pos < end) {
        while (pos < end && buf[pos] <= 0x20) pos++;
        const nameStart = pos;
        while (pos < end && buf[pos] !== 0x3d && buf[pos] > 0x20) pos++;
        const nameEnd = pos;
        while (pos < end && buf[pos] <= 0x20) pos++;
        if (pos >= end || buf[pos] !== 0x3d) {
            pos++;
            continue;
        }
        pos++;
        while (pos < end && buf[pos] <= 0x20) pos++;
        if (pos >= end || (buf[pos] !== 0x22 && buf[pos] !== 0x27)) continue;
        const quote = buf[pos++];
        const valueStart = pos;
        while (pos < end && buf[pos] !== quote) pos++;
        const valueEnd = pos;
        if (nameEnd - nameStart === name.length) {
            let match = true;
            for (let i = 0; i < name.length; i++) {
                if (buf[nameStart + i] !== name.charCodeAt(i)) { match = false; break; }
            }
            if (match) return buf.toString('utf8', valueStart, valueEnd);
        }
        if (pos < end) pos++;
    }
    return null;
}

/** ASCII unsigned integer in `[start, end)`, or -1. */
function parseUint(buf: Buffer, start: number, end: number): number {
    let value = 0;
    if (start >= end) return -1;
    for (let i = start; i < end; i++) {
        const c = buf[i];
        if (c < 0x30 || c > 0x39) return -1;
        value = value * 10 + (c - 0x30);
    }
    return value;
}

/**
 * Parse a plain ASCII decimal number without allocating a string.
 *
 * Integers and decimals with at most 15 significant digits are exact (the
 * mantissa fits in 2^53 and the single division by 10^n is correctly
 * rounded), so the result matches `parseFloat`. Everything else (exponents,
 * longer mantissas, surrounding whitespace) falls back to `parseFloat`.
 */
function parseNumberBytes(buf: Buffer, start: number, end: number): number {
    let i = start;
    let negative = false;
    if (i < end) {
        const c = buf[i];
        if (c === 0x2d) {
            negative = true;
            i++;
        } else if (c === 0x2b) {
            i++;
        }
    }
    let mantissa = 0;
    let digits = 0;
    let decimals = 0;
    let sawDot = false;
    let sawDigit = false;
    for (; i < end; i++) {
        const c = buf[i];
        if (c >= 0x30 && c <= 0x39) {
            if (digits >= 15) return parseFloat(buf.toString('utf8', start, end));
            mantissa = mantissa * 10 + (c - 0x30);
            digits++;
            sawDigit = true;
            if (sawDot) decimals++;
        } else if (c === 0x2e && !sawDot) {
            sawDot = true;
        } else {
            return parseFloat(buf.toString('utf8', start, end));
        }
    }
    if (!sawDigit) return NaN;
    const value = decimals === 0 ? mantissa : mantissa / Math.pow(10, decimals);
    return negative ? -value : value;
}

const CELL_TYPE_NUM = 0;
const CELL_TYPE_SHARED = 1;
const CELL_TYPE_BOOL = 2;
const CELL_TYPE_ERROR = 3;
const CELL_TYPE_DATE = 4;
const CELL_TYPE_INLINE = 5;
const CELL_TYPE_STR = 6;

/** Decode a cell `t` attribute without allocating a string. */
function parseCellType(buf: Buffer, start: number, end: number): number {
    const length = end - start;
    if (length === 1) {
        const c = buf[start];
        if (c === 0x73) return CELL_TYPE_SHARED; // s
        if (c === 0x62) return CELL_TYPE_BOOL;   // b
        if (c === 0x65) return CELL_TYPE_ERROR;  // e
        if (c === 0x64) return CELL_TYPE_DATE;   // d
        return CELL_TYPE_NUM;
    }
    if (length === 3 && buf[start] === 0x73 && buf[start + 1] === 0x74 && buf[start + 2] === 0x72) {
        return CELL_TYPE_STR; // str
    }
    if (length === 9 &&
        buf[start] === 0x69 && buf[start + 1] === 0x6e && buf[start + 2] === 0x6c &&
        buf[start + 3] === 0x69 && buf[start + 4] === 0x6e && buf[start + 5] === 0x65 &&
        buf[start + 6] === 0x53 && buf[start + 7] === 0x74 && buf[start + 8] === 0x72) {
        return CELL_TYPE_INLINE; // inlineStr
    }
    return CELL_TYPE_NUM;
}

/** 0-based column from an `r` attribute value, decoded in place. */
function parseCellColumnBytes(buf: Buffer, start: number, end: number): number {
    let value = 0;
    let letters = 0;
    for (let i = start; i < end; i++) {
        const c = buf[i];
        if (c >= 65 && c <= 90) {
            value = value * 26 + (c - 64);
            letters++;
        } else if (c >= 97 && c <= 122) {
            value = value * 26 + (c - 96);
            letters++;
        } else {
            break;
        }
    }
    return letters > 0 && letters <= 3 ? value - 1 : -1;
}

interface CellAttributes {
    col: number;
    type: number;
    style: number;
}

/** Parse `r`, `t` and `s` from a `<c …>` tag in one pass. */
function parseCellAttributes(buf: Buffer, start: number, end: number): CellAttributes {
    let col = -1;
    let type = CELL_TYPE_NUM;
    let style = 0;

    let pos = start + 1;
    // Skip the element name.
    while (pos < end && buf[pos] > 0x20 && buf[pos] !== 0x3e && buf[pos] !== 0x2f) pos++;

    while (pos < end) {
        while (pos < end && buf[pos] <= 0x20) pos++;
        if (pos >= end) break;
        const nameStart = pos;
        while (pos < end && buf[pos] !== 0x3d && buf[pos] > 0x20) pos++;
        const nameEnd = pos;
        while (pos < end && buf[pos] <= 0x20) pos++;
        if (pos >= end || buf[pos] !== 0x3d) {
            if (pos === nameStart) pos++;
            continue;
        }
        pos++;
        while (pos < end && buf[pos] <= 0x20) pos++;
        if (pos >= end || (buf[pos] !== 0x22 && buf[pos] !== 0x27)) continue;
        const quote = buf[pos++];
        const valueStart = pos;
        while (pos < end && buf[pos] !== quote) pos++;
        const valueEnd = pos;

        if (nameEnd - nameStart === 1) {
            const name = buf[nameStart];
            if (name === 0x72) {
                col = parseCellColumnBytes(buf, valueStart, valueEnd);
            } else if (name === 0x74) {
                type = parseCellType(buf, valueStart, valueEnd);
            } else if (name === 0x73) {
                const parsed = parseUint(buf, valueStart, valueEnd);
                if (parsed >= 0) style = parsed;
            }
        }
        if (pos < end) pos++;
    }

    return { col, type, style };
}

export class XlsxReader extends ExcelReaderAbstract {
    private zipfile: ZipFile | null = null;
    /** Keeps the yauzl source buffer alive when opened from memory. */
    private _zipBuffer: Buffer | null = null;
    private sharedStrings: string[] = [];
    private sheetNames: string[] = [];
    private sheets: SheetInfo[] = [];
    private entries: Map<string, Entry> = new Map();

    private _currentSheetIndex: number = -1;
    /** Buffer-based fast path for the current sheet. */
    private _sheetBuffer: Buffer | null = null;
    /** String fallback for unusual markup (prefixed tags, etc.). */
    private _sheetXml: string | null = null;
    private _xmlPos: number = 0;
    private _xmlEnd: number = 0;

    private _currentRow: CellValue[] = [];
    private _currentFormulas: Array<string | null> = [];
    private _currentStyleIndices: number[] = [];
    private _currentSharedIndices: number[] = [];
    private _currentRichTexts: Array<RichText | null> = [];
    private _expectedRowNum: number = 1;
    private _pendingEmptyCount: number = 0;

    private _xfStyles: Array<CellStyle | null> = [];
    private _sharedRichTexts: Array<RichText | null> = [];

    xfIdToNumFmtId: number[] = [];
    customDateFormats: Set<number> = new Set();
    /** One flag per xf id: 1 when the style's number format is a date. */
    private _dateXf: Uint8Array = new Uint8Array(0);

    constructor() {
        super();
    }

    async open(path: string, readSharedStrings: boolean = true): Promise<void> {
        await this._openWith(
            (callback) => yauzl.open(path, { lazyEntries: true, autoClose: false }, callback),
            readSharedStrings,
        );
    }

    /** Open a workbook already held in memory (ZIP bytes). */
    async openBuffer(buffer: Buffer, readSharedStrings: boolean = true): Promise<void> {
        if (!Buffer.isBuffer(buffer)) {
            throw new TypeError('XlsxReader.openBuffer expects a Buffer');
        }
        this._zipBuffer = buffer;
        await this._openWith(
            (callback) => yauzl.fromBuffer(buffer, { lazyEntries: true, autoClose: false }, callback),
            readSharedStrings,
        );
    }

    private async _openWith(
        opener: (callback: (err: Error | null, zipfile?: ZipFile) => void) => void,
        readSharedStrings: boolean,
    ): Promise<void> {
        await this.close();
        this._resetState();
        return new Promise((resolve, reject) => {
            opener((err, zipfile) => {
                if (err || !zipfile) return reject(err ?? new Error('XlsxReader: could not open ZIP'));
                this.zipfile = zipfile;
                this.entries = new Map();

                zipfile.on('entry', (entry: Entry) => {
                    this.entries.set(entry.fileName, entry);
                    zipfile.readEntry();
                });

                zipfile.on('end', async () => {
                    try {
                        await this._loadWorkbookStructure(readSharedStrings);
                        resolve();
                    } catch (e) {
                        reject(e);
                    }
                });
                zipfile.readEntry();
            });
        });
    }

    private _resetState(): void {
        this.sheetNames = [];
        this.sheets = [];
        this.sharedStrings = [];
        this.xfIdToNumFmtId = [];
        this.customDateFormats = new Set();
        this._dateXf = new Uint8Array(0);
        this.fieldCount = 0;
        this.resultsCount = 0;
        this.actualSheetName = '';
        this._currentSheetIndex = -1;
        this._sheetBuffer = null;
        this._sheetXml = null;
        this._xmlPos = 0;
        this._xmlEnd = 0;
        this._resetCurrentRow();
        this._currentFormulas = [];
        this._currentStyleIndices = [];
        this._currentSharedIndices = [];
        this._xfStyles = [];
        this._sharedRichTexts = [];
        this._expectedRowNum = 1;
        this._pendingEmptyCount = 0;
        this._oaEpoch = Date.UTC(1899, 11, 30);
    }

    private async _loadWorkbookStructure(readSharedStrings: boolean): Promise<void> {
        const wbRelsContent = await this._readZipEntryContent('xl/_rels/workbook.xml.rels');
        const rIdToTarget: Record<string, string> = {};
        if (wbRelsContent) {
            const relTagRe = /<Relationship\b[^>]*>/g;
            let tagMatch;
            while ((tagMatch = relTagRe.exec(wbRelsContent)) !== null) {
                const attrs = parseTagAttributes(tagMatch[0]);
                const id = attrs['Id'] ?? attrs['id'];
                const target = attrs['Target'] ?? attrs['target'];
                if (id !== undefined && target !== undefined) rIdToTarget[id] = target;
            }
        }

        const wbContent = await this._readZipEntryContent('xl/workbook.xml');
        if (wbContent) {
            // Order-independent: parse each <sheet> tag then its attributes.
            const sheetTagRe = /<sheet\b[^>]*>/g;
            let tagMatch;
            while ((tagMatch = sheetTagRe.exec(wbContent)) !== null) {
                const attrs = parseTagAttributes(tagMatch[0]);
                const rawName = attrs['name'];
                const sheetId = attrs['sheetId'] ?? '';
                const rId = attrs['r:id'] ?? attrs['id'] ?? attrs['rId'];
                if (rawName === undefined || rId === undefined) continue;
                const name = unescapeXml(rawName);
                const target = rIdToTarget[rId];
                if (!target) continue;
                this.sheetNames.push(name);
                this.sheets.push({ name, sheetId, rId, path: resolveRelTarget(target) });
            }
            const wbPrMatch = /<workbookPr\b[^>]*>/i.exec(wbContent);
            if (wbPrMatch) {
                const wbPrAttrs = parseTagAttributes(wbPrMatch[0]);
                const v = (wbPrAttrs['date1904'] ?? '').trim().toLowerCase();
                this._oaEpoch = (v === '1' || v === 'true') ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
            }
        }

        if (readSharedStrings) {
            const ssEntry = this.entries.get('xl/sharedStrings.xml');
            if (ssEntry) {
                const ssBuffer = await this._readZipEntryBuffer(ssEntry);
                this.sharedStrings = parseSharedStringsBuffer(ssBuffer);
                if (ssBuffer.includes('<r>') || ssBuffer.includes('<r ')) {
                    this._sharedRichTexts = parseXlsxRichTexts(ssBuffer.toString('utf8'));
                }
            }
        }

        const stylesEntry = this.entries.get('xl/styles.xml');
        if (stylesEntry) {
            const stylesBuffer = await this._readZipEntryBuffer(stylesEntry);
            this._parseStyles(stylesBuffer.toString('utf8'));
        }

        this.resultsCount = this.sheets.length;
        this._currentSheetIndex = -1;
    }

    async close(): Promise<void> {
        if (this.zipfile) {
            this.zipfile.close();
            this.zipfile = null;
        }
        this._zipBuffer = null;
    }

    private async _readZipEntryContent(path: string): Promise<string | null> {
        const entry = this.entries.get(path);
        if (!entry) return null;
        return (await this._readZipEntryBuffer(entry)).toString('utf8');
    }

    /**
     * Read one ZIP member into a Buffer. The uncompressed size from the
     * central directory is used to allocate once instead of concatenating
     * stream chunks.
     */
    private _readZipEntryBuffer(entry: Entry): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            this.zipfile!.openReadStream(entry, (err, readStream) => {
                if (err || !readStream) return reject(err ?? new Error('XlsxReader: cannot read ZIP entry'));
                // Cap the one-shot allocation: the size comes from the (untrusted)
                // central directory, and larger entries fall back to chunk concat.
                const MAX_PREALLOC = 64 * 1024 * 1024;
                const expected = entry.uncompressedSize;
                const prealloc = expected > 0 && expected <= MAX_PREALLOC
                    ? Buffer.allocUnsafe(expected)
                    : null;
                let written = 0;
                let overflowed = false;
                const overflow: Buffer[] = [];
                readStream.on('data', (chunk: Buffer) => {
                    if (!overflowed && prealloc && written + chunk.length <= prealloc.length) {
                        chunk.copy(prealloc, written);
                        written += chunk.length;
                    } else {
                        // Once a chunk does not fit, keep every later chunk in
                        // order; the preallocated tail must not be filled.
                        overflowed = true;
                        overflow.push(chunk);
                    }
                });
                readStream.on('end', () => {
                    if (prealloc && !overflowed) {
                        resolve(prealloc.subarray(0, written));
                        return;
                    }
                    const parts: Buffer[] = [];
                    if (prealloc && written > 0) parts.push(prealloc.subarray(0, written));
                    if (overflow.length > 0) parts.push(...overflow);
                    resolve(Buffer.concat(parts));
                });
                readStream.on('error', reject);
            });
        });
    }

    private _parseSharedStrings(xml: string): void {
        this.sharedStrings = parseSharedStringsXml(xml);
    }

    private _parseStyles(xml: string): void {
        const numFmtRegex = /<numFmt\s+[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
        let match;
        this.customDateFormats = new Set();

        while ((match = numFmtRegex.exec(xml)) !== null) {
            const id = parseInt(match[1]);
            const code = match[2].toLowerCase();
            if (code.includes('yy') || code.includes('mm') || code.includes('dd') || code.includes('h:mm')) {
                this.customDateFormats.add(id);
            }
        }

        this.xfIdToNumFmtId = [];
        const cellXfsStart = xml.indexOf('<cellXfs');
        if (cellXfsStart !== -1) {
            const cellXfsEnd = xml.indexOf('</cellXfs>', cellXfsStart);
            if (cellXfsEnd !== -1) {
                const cellXfsContent = xml.substring(cellXfsStart, cellXfsEnd);
                const xfRegex = /<xf\s+[^>]*numFmtId="(\d+)"/g;
                while ((match = xfRegex.exec(cellXfsContent)) !== null) {
                    this.xfIdToNumFmtId.push(parseInt(match[1]));
                }
            }
        }

        this._dateXf = new Uint8Array(this.xfIdToNumFmtId.length);
        for (let i = 0; i < this.xfIdToNumFmtId.length; i++) {
            this._dateXf[i] = this._isDateNumFmt(this.xfIdToNumFmtId[i]) ? 1 : 0;
        }

        this._xfStyles = parseXlsxStyles(xml).cellStyles;
    }

    private _isDateNumFmt(numFmtId: number): boolean {
        return (numFmtId >= 14 && numFmtId <= 22) ||
            (numFmtId >= 45 && numFmtId <= 47) ||
            this.customDateFormats.has(numFmtId);
    }

    private _isDateStyle(styleIndex: number): boolean {
        return styleIndex > 0 && styleIndex < this._dateXf.length && this._dateXf[styleIndex] !== 0;
    }

    getSheetNames(): string[] {
        return this.sheetNames;
    }

    /** Advance to the next row. Always use `await`. */
    async read(): Promise<boolean> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            await this._initSheet(0);
        }
        return this.readRow() !== null;
    }

    /**
     * Synchronous row access for batch consumers. Returns the internal row
     * array, valid until the next call; `readAll`/`rows` return copies.
     * Select a sheet (or call `read()` once) before using it.
     */
    readRow(): CellValue[] | null {
        if (this._currentSheetIndex === -1) {
            throw new Error('XlsxReader: select a sheet before readRow()');
        }
        const hasRow = this._sheetBuffer !== null ? this._readNextRowBuffer() : this._readNextRowString();
        return hasRow ? this._currentRow : null;
    }

    /** Collect every remaining row of the current sheet. */
    async readAll(): Promise<CellValue[][]> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            await this._initSheet(0);
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
            await this._initSheet(0);
        }
        while (true) {
            const row = this.readRow();
            if (row === null) return;
            yield row.slice();
        }
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
        await this._initSheet(index);
    }

    async nextResult(): Promise<boolean> {
        const next = this._currentSheetIndex + 1;
        if (next >= this.sheets.length) return false;
        this._currentSheetIndex = next;
        await this._initSheet(next);
        return true;
    }

    private async _initSheet(index: number): Promise<void> {
        if (index >= this.sheets.length) return;
        const sheet = this.sheets[index];
        this.actualSheetName = sheet.name;

        const entry = this.entries.get(sheet.path);
        if (!entry) {
            this._sheetBuffer = null;
            this._sheetXml = '';
            this._xmlPos = 0;
            this._xmlEnd = 0;
            this._expectedRowNum = 1;
            this._pendingEmptyCount = 0;
            this._resetCurrentRow();
            this.fieldCount = 0;
            return;
        }

        const buffer = await this._readZipEntryBuffer(entry);
        this._expectedRowNum = 1;
        this._pendingEmptyCount = 0;
        this._resetCurrentRow();
        this.fieldCount = 0;

        this._sheetBuffer = null;
        this._sheetXml = null;
        if (this._initBufferSheet(buffer)) return;

        // Unusual markup (prefixed names, etc.): use the string parser.
        this._sheetXml = buffer.toString('utf8');
        this._xmlPos = 0;
        this._xmlEnd = this._sheetXml.length;
        this._initStringSheetRange();
    }

    /** Set `_xmlPos/_xmlEnd` for the buffer fast path; false when unsupported. */
    private _initBufferSheet(buffer: Buffer): boolean {
        let contentStart = -1;
        let contentEnd = -1;
        const simpleStart = buffer.indexOf('<sheetData>');
        if (simpleStart !== -1) {
            contentStart = simpleStart + '<sheetData>'.length;
            contentEnd = buffer.indexOf('</sheetData>', contentStart);
            if (contentEnd === -1) return false;
        } else {
            // Detect a prefixed or attribute-bearing sheetData; the generic
            // string path understands those.
            const head = buffer.toString('latin1', 0, Math.min(buffer.length, 8192));
            if (/<[A-Za-z_][\w.-]*:sheetData\b/i.test(head)) return false;
            const open = buffer.indexOf('<sheetData');
            if (open === -1) {
                // No sheetData at all: empty sheet.
                this._sheetBuffer = buffer;
                this._xmlPos = 0;
                this._xmlEnd = 0;
                return true;
            }
            const end = tagEndIndex(buffer, open, buffer.length);
            if (end === -1) return false;
            if (buffer[end - 1] === 0x2f) {
                this._sheetBuffer = buffer;
                this._xmlPos = end + 1;
                this._xmlEnd = end + 1;
                return true;
            }
            contentStart = end + 1;
            contentEnd = buffer.indexOf('</sheetData>', contentStart);
            if (contentEnd === -1) return false;
        }
        this._sheetBuffer = buffer;
        this._xmlPos = contentStart;
        this._xmlEnd = contentEnd;
        return true;
    }

    private _initStringSheetRange(): void {
        if (!this._sheetXml) return;
        const sd = this._sheetXml.indexOf('<sheetData>');
        if (sd !== -1) {
            this._xmlPos = sd + 11;
            const se = this._sheetXml.indexOf('</sheetData>', this._xmlPos);
            if (se !== -1) this._xmlEnd = se;
            return;
        }
        const m = /<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/i.exec(this._sheetXml);
        if (m && m.index !== undefined) {
            this._xmlPos = m.index + m[0].length;
            if (/\/>\s*$/.test(m[0])) {
                this._xmlEnd = this._xmlPos;
            } else {
                const se = this._sheetXml.indexOf('</sheetData>', this._xmlPos);
                if (se !== -1) this._xmlEnd = se;
            }
        }
    }

    // ------------------------------------------------------------------
    // Buffer fast path
    // ------------------------------------------------------------------

    private _readNextRowBuffer(): boolean {
        const buffer = this._sheetBuffer!;
        if (this._pendingEmptyCount > 0) {
            this._pendingEmptyCount--;
            this._resetCurrentRow();
            this._expectedRowNum++;
            return true;
        }

        const rowStart = this._findRowTagStartBuffer(buffer);
        if (rowStart === -1) return false;
        const rowTagEnd = tagEndIndex(buffer, rowStart, this._xmlEnd);
        if (rowTagEnd === -1) return false;
        const isSelfClosing = buffer[rowTagEnd - 1] === 0x2f;
        const rRaw = attributeValue(buffer, rowStart, rowTagEnd, 'r');
        let rowNum = this._expectedRowNum;
        if (rRaw !== null) {
            const parsed = parseInt(rRaw, 10);
            if (Number.isFinite(parsed) && parsed >= 1 && parsed <= MAX_ROWS) rowNum = parsed;
        }

        if (rowNum > this._expectedRowNum) {
            const gap = rowNum - this._expectedRowNum;
            this._pendingEmptyCount = gap - 1;
            this._resetCurrentRow();
            this._expectedRowNum++;
            return true;
        }
        if (rowNum < this._expectedRowNum) {
            if (isSelfClosing) {
                this._xmlPos = rowTagEnd + 1;
                this._resetCurrentRow();
                return true;
            }
            const rowEnd = buffer.indexOf('</row>', rowStart);
            if (rowEnd === -1) return false;
            this._xmlPos = rowEnd + 6;
            this._parseRowBuffer(rowStart, rowEnd);
            this._expectedRowNum = Math.max(this._expectedRowNum + 1, rowNum + 1);
            return true;
        }

        if (isSelfClosing) {
            this._xmlPos = rowTagEnd + 1;
            this._resetCurrentRow();
            this._expectedRowNum = rowNum + 1;
            return true;
        }

        const rowEnd = buffer.indexOf('</row>', rowStart);
        if (rowEnd === -1) return false;
        this._xmlPos = rowEnd + 6;
        this._parseRowBuffer(rowStart, rowEnd);
        this._expectedRowNum = rowNum + 1;
        return true;
    }

    private _findRowTagStartBuffer(buffer: Buffer): number {
        let pos = this._xmlPos;
        while (pos < this._xmlEnd) {
            const idx = buffer.indexOf('<row', pos);
            if (idx === -1 || idx >= this._xmlEnd) return -1;
            const next = buffer[idx + 4];
            if (next === 0x3e || next === 0x2f || next === 0x20 ||
                next === 0x09 || next === 0x0a || next === 0x0d) {
                return idx;
            }
            pos = idx + 4;
        }
        return -1;
    }

    private _parseRowBuffer(rowStart: number, rowEnd: number): void {
        const buffer = this._sheetBuffer!;
        // Reuse the row arrays: readRow() documents that its result is only
        // valid until the next call, and reuse avoids five allocations per row.
        const row = this._currentRow;
        const styles = this._currentStyleIndices;
        const formulas = this._currentFormulas;
        const sharedIndices = this._currentSharedIndices;
        const richTexts = this._currentRichTexts;
        row.length = 0;
        styles.length = 0;
        formulas.length = 0;
        sharedIndices.length = 0;
        richTexts.length = 0;

        let pos = tagEndIndex(buffer, rowStart, rowEnd) + 1;
        let nextCol = 0;

        while (pos < rowEnd) {
            const cStart = indexOfTag(buffer, 0x63, pos, rowEnd);
            if (cStart === -1) break;
            const afterName = buffer[cStart + 2];
            if (afterName !== 0x3e && afterName !== 0x2f && afterName !== 0x20 &&
                afterName !== 0x09 && afterName !== 0x0a && afterName !== 0x0d) {
                pos = cStart + 2;
                continue;
            }
            const tagEnd = tagEndIndex(buffer, cStart, rowEnd);
            if (tagEnd === -1) break;

            const attrs = parseCellAttributes(buffer, cStart, tagEnd);
            const type = attrs.type;
            const styleIndex = attrs.style;

            let finalVal: CellValue = null;
            let formula: string | null = null;
            let sharedIndex = -1;
            let richText: RichText | null = null;
            if (buffer[tagEnd - 1] === 0x2f) {
                pos = tagEnd + 1;
            } else {
                const contentStart = tagEnd + 1;
                const cEnd = indexOfCloseTag(buffer, 0x63, contentStart, rowEnd);
                const contentEnd = cEnd === -1 ? rowEnd : cEnd;

                // ECMA-376 puts <f> first when present, so this avoids scanning
                // every cell for a formula that is almost never there.
                if (buffer[contentStart] === BYTE_LT && buffer[contentStart + 1] === 0x66) {
                    formula = this._extractFormulaBuffer(buffer, contentStart, contentEnd);
                }

                const vOpen = indexOfTag(buffer, 0x76, contentStart, contentEnd);
                if (vOpen !== -1) {
                    const vTagEnd = tagEndIndex(buffer, vOpen, contentEnd);
                    if (vTagEnd !== -1) {
                        const valueStart = vTagEnd + 1;
                        // `<` never appears literally inside <v>; the next one
                        // closes the element.
                        const vEnd = indexOfByte(buffer, BYTE_LT, valueStart, contentEnd);
                        if (vEnd !== -1) {
                            finalVal = this._decodeCellValue(
                                buffer, valueStart, vEnd, type, styleIndex,
                            );
                            if (type === CELL_TYPE_SHARED) {
                                sharedIndex = parseUint(buffer, valueStart, vEnd);
                            }
                        }
                    }
                } else if (type === CELL_TYPE_INLINE) {
                    finalVal = this._decodeInlineBuffer(buffer, contentStart, contentEnd);
                    // Skip the rich-text regexes for plain inline strings.
                    if (indexOfTag(buffer, 0x72, contentStart, contentEnd) !== -1) {
                        richText = parseXlsxInlineRichText(
                            buffer.toString('utf8', contentStart, contentEnd),
                        );
                    }
                }
                pos = contentEnd + 4;
            }

            const normalizedCol = attrs.col === -1 ? nextCol : attrs.col;
            nextCol = normalizedCol + 1;
            row[normalizedCol] = finalVal;
            styles[normalizedCol] = styleIndex;
            formulas[normalizedCol] = formula;
            sharedIndices[normalizedCol] = sharedIndex;
            richTexts[normalizedCol] = richText;
            if (normalizedCol >= this.fieldCount) this.fieldCount = normalizedCol + 1;
        }
    }

    /** Concatenated `<t>` runs of an inline-string cell body. */
    private _decodeInlineBuffer(buffer: Buffer, start: number, end: number): CellValue {
        let value = '';
        let sawText = false;
        let pos = start;
        while (true) {
            const tStart = indexOfTag(buffer, 0x74, pos, end);
            if (tStart === -1) break;
            const tTagEnd = tagEndIndex(buffer, tStart, end);
            if (tTagEnd === -1) break;
            const tContentStart = tTagEnd + 1;
            const tEnd = indexOfCloseTag(buffer, 0x74, tContentStart, end);
            if (tEnd === -1) break;
            value += buffer.toString('utf8', tContentStart, tEnd);
            sawText = true;
            pos = tEnd + 4;
        }
        return sawText ? unescapeXmlText(value) : null;
    }

    private _extractFormulaBuffer(buffer: Buffer, start: number, end: number): string | null {
        // The caller already verified that the cell body starts with `<f`.
        const fTagEnd = tagEndIndex(buffer, start, end);
        if (fTagEnd === -1 || buffer[fTagEnd - 1] === 0x2f) return null;
        const fEnd = indexOfCloseTag(buffer, 0x66, fTagEnd + 1, end);
        if (fEnd === -1) return null;
        return unescapeXml(buffer.toString('utf8', fTagEnd + 1, fEnd));
    }

    private _decodeCellValue(buffer: Buffer, start: number, end: number, type: number, styleIndex: number): CellValue {
        if (type === CELL_TYPE_SHARED) {
            const index = parseUint(buffer, start, end);
            return index >= 0 ? (this.sharedStrings[index] ?? null) : null;
        }
        if (type === CELL_TYPE_NUM) {
            const value = parseNumberBytes(buffer, start, end);
            if (this._isDateStyle(styleIndex)) {
                try {
                    return this.getDateTimeFromOaDate(value);
                } catch {
                    return value;
                }
            }
            return value;
        }
        const text = buffer.toString('utf8', start, end);
        if (type === CELL_TYPE_BOOL) {
            return text === '1' || text === 'true';
        }
        if (type === CELL_TYPE_INLINE || type === CELL_TYPE_STR) {
            return text.indexOf('&') === -1 && text.indexOf('_x') === -1 ? text : unescapeXmlText(text);
        }
        if (type === CELL_TYPE_ERROR) {
            return unescapeXml(text);
        }
        if (type === CELL_TYPE_DATE) {
            return parseIsoDateUtc(text) ?? unescapeXmlText(text);
        }
        return parseFloat(text);
    }

    // ------------------------------------------------------------------
    // String fallback (unusual markup)
    // ------------------------------------------------------------------

    private _readNextRowString(): boolean {
        if (this._sheetXml === null) return false;
        if (this._pendingEmptyCount > 0) {
            this._pendingEmptyCount--;
            this._resetCurrentRow();
            this._expectedRowNum++;
            return true;
        }

        const rowStart = this._findRowTagStartString(this._xmlPos);
        if (rowStart === -1) return false;

        const rowTagEnd = this._sheetXml.indexOf('>', rowStart);
        if (rowTagEnd === -1) return false;
        const rowTag = this._sheetXml.substring(rowStart, rowTagEnd + 1);
        const isSelfClosing = this._sheetXml.charCodeAt(rowTagEnd - 1) === 47;
        const rowAttrs = parseTagAttributes(rowTag);
        const rRaw = rowAttrs['r'];
        let rowNum = this._expectedRowNum;
        if (rRaw !== undefined) {
            const parsed = parseInt(rRaw, 10);
            if (Number.isFinite(parsed) && parsed >= 1 && parsed <= MAX_ROWS) rowNum = parsed;
        }

        if (rowNum > this._expectedRowNum) {
            const gap = rowNum - this._expectedRowNum;
            this._pendingEmptyCount = gap - 1;
            this._resetCurrentRow();
            this._expectedRowNum++;
            return true;
        }
        if (rowNum < this._expectedRowNum) {
            if (isSelfClosing) {
                this._xmlPos = rowTagEnd + 1;
                this._resetCurrentRow();
                return true;
            }
            const rowEnd = this._sheetXml.indexOf('</row>', rowStart);
            if (rowEnd === -1) return false;
            this._xmlPos = rowEnd + 6;
            this._parseRowString(rowStart, rowEnd);
            this._expectedRowNum = Math.max(this._expectedRowNum + 1, rowNum + 1);
            return true;
        }

        if (isSelfClosing) {
            this._xmlPos = rowTagEnd + 1;
            this._resetCurrentRow();
            this._expectedRowNum = rowNum + 1;
            return true;
        }

        const rowEnd = this._sheetXml.indexOf('</row>', rowStart);
        if (rowEnd === -1) return false;
        this._xmlPos = rowEnd + 6;
        this._parseRowString(rowStart, rowEnd);
        this._expectedRowNum = rowNum + 1;
        return true;
    }

    private _findRowTagStartString(from: number): number {
        const xml = this._sheetXml!;
        let pos = from;
        while (pos < this._xmlEnd) {
            const idx = xml.indexOf('<row', pos);
            if (idx === -1 || idx >= this._xmlEnd) return -1;
            const next = xml.charCodeAt(idx + 4);
            if (next === 0x3e || next === 0x2f || next === 0x20 ||
                next === 0x09 || next === 0x0a || next === 0x0d) {
                return idx;
            }
            pos = idx + 4;
        }
        return -1;
    }

    private _parseRowString(rowStart: number, rowEnd: number): void {
        this._resetCurrentRow();
        const styles: number[] = [];
        const formulas: Array<string | null> = [];
        const sharedIndices: number[] = [];
        const richTexts: Array<RichText | null> = [];
        const xml = this._sheetXml!;
        let pos = xml.indexOf('>', rowStart) + 1;
        let nextCol = 0;

        while (pos < rowEnd) {
            const cStart = stringIndexOfWithin(xml, '<c', pos, rowEnd);
            if (cStart === -1) break;
            const nextTagClose = stringIndexOfWithin(xml, '>', cStart, rowEnd);
            if (nextTagClose === -1) break;
            let colIndex = -1;
            let type = 'n';
            let styleIndex = 0;

            let currentIdx = cStart + 2;
            while (currentIdx < nextTagClose) {
                while (xml.charCodeAt(currentIdx) <= 32 && currentIdx < nextTagClose) currentIdx++;
                if (currentIdx >= nextTagClose) break;
                const keyStart = currentIdx;
                while (xml.charCodeAt(currentIdx) !== 61 && currentIdx < nextTagClose) currentIdx++;
                const key = xml.substring(keyStart, currentIdx);
                currentIdx++;
                if (xml.charCodeAt(currentIdx) === 34) {
                    currentIdx++;
                    const valStart = currentIdx;
                    while (xml.charCodeAt(currentIdx) !== 34 && currentIdx < nextTagClose) currentIdx++;
                    const val = xml.substring(valStart, currentIdx);
                    currentIdx++;
                    if (key === 'r') {
                        let letterLen = 0;
                        while (letterLen < val.length && val.charCodeAt(letterLen) >= 65) letterLen++;
                        colIndex = columnLetterToIndex(val.substring(0, letterLen));
                    } else if (key === 't') {
                        type = val;
                    } else if (key === 's') {
                        styleIndex = parseInt(val, 10);
                    }
                }
            }

            let val: string | null = null;
            let inlineVals: string[] | null = null;
            let richText: RichText | null = null;
            let cellContentEnd = rowEnd;
            if (xml.charCodeAt(nextTagClose - 1) === 47) {
                // A self-closing cell has no body; keep the scan range empty so
                // the formula/rich-text extraction below does not read later cells.
                cellContentEnd = nextTagClose;
                pos = nextTagClose + 1;
            } else {
                const cEnd = stringIndexOfWithin(xml, '</c>', nextTagClose, rowEnd);
                cellContentEnd = cEnd === -1 ? rowEnd : cEnd;
                const vStart = stringIndexOfWithin(xml, '<v>', nextTagClose, cellContentEnd);
                if (vStart !== -1) {
                    const vInnerStart = vStart + 3;
                    const vEnd = stringIndexOfWithin(xml, '</v>', vInnerStart, cellContentEnd);
                    if (vEnd !== -1) val = xml.substring(vInnerStart, vEnd);
                } else if (type === 'inlineStr') {
                    inlineVals = [];
                    let tSearch = nextTagClose;
                    while (true) {
                        const tStart = stringIndexOfWithin(xml, '<t', tSearch, cellContentEnd);
                        if (tStart === -1) break;
                        const tClose = stringIndexOfWithin(xml, '>', tStart, cellContentEnd);
                        if (tClose === -1) break;
                        const tContentStart = tClose + 1;
                        const tEnd = stringIndexOfWithin(xml, '</t>', tContentStart, cellContentEnd);
                        if (tEnd === -1) break;
                        inlineVals.push(xml.substring(tContentStart, tEnd));
                        tSearch = tEnd + 4;
                    }
                    if (inlineVals.length === 0) inlineVals = null;
                }
                pos = cellContentEnd + 4;
            }

            let finalVal: CellValue = null;
            if (inlineVals !== null) {
                finalVal = unescapeXmlText(inlineVals.join(''));
            } else if (val !== null) {
                if (type === 's') {
                    finalVal = this.sharedStrings[parseInt(val, 10)] ?? null;
                } else if (type === 'b') {
                    finalVal = (val === '1' || val === 'true');
                } else if (type === 'inlineStr' || type === 'str') {
                    finalVal = unescapeXmlText(val);
                } else if (type === 'e') {
                    finalVal = unescapeXml(val);
                } else if (type === 'd') {
                    finalVal = parseIsoDateUtc(val) ?? unescapeXmlText(val);
                } else {
                    finalVal = parseFloat(val);
                    if (this._isDateStyle(styleIndex)) {
                        try {
                            finalVal = this.getDateTimeFromOaDate(finalVal);
                        } catch {
                            // keep the serial
                        }
                    }
                }
            }

            if (colIndex === -1) colIndex = nextCol;
            nextCol = colIndex + 1;
            this._currentRow[colIndex] = finalVal;
            styles[colIndex] = styleIndex;
            formulas[colIndex] = this._extractFormulaString(xml, nextTagClose, cellContentEnd);
            if (type === 'inlineStr' && cellContentEnd > nextTagClose) {
                richText = parseXlsxInlineRichText(xml.substring(nextTagClose + 1, cellContentEnd));
            }
            richTexts[colIndex] = richText;
            sharedIndices[colIndex] = type === 's' && val !== null
                ? (Number.isFinite(parseInt(val, 10)) ? parseInt(val, 10) : -1)
                : -1;
            if (colIndex >= this.fieldCount) this.fieldCount = colIndex + 1;
        }
        this._currentStyleIndices = styles;
        this._currentFormulas = formulas;
        this._currentSharedIndices = sharedIndices;
        this._currentRichTexts = richTexts;
    }

    private _extractFormulaString(xml: string, start: number, end: number): string | null {
        if (end <= start) return null;
        const fStart = stringIndexOfWithin(xml, '<f', start, end);
        if (fStart === -1) return null;
        const fTagEnd = stringIndexOfWithin(xml, '>', fStart, end);
        if (fTagEnd === -1 || xml.charCodeAt(fTagEnd - 1) === 47) return null;
        const fEnd = stringIndexOfWithin(xml, '</f>', fTagEnd + 1, end);
        if (fEnd === -1) return null;
        return unescapeXml(xml.substring(fTagEnd + 1, fEnd));
    }

    /** Style of the current row cell at column `i`, or null for the default style. */
    getStyle(i: number): CellStyle | null {
        const xf = this._currentStyleIndices[i];
        if (xf === undefined) return null;
        return this._xfStyles[xf] ?? null;
    }

    /** Formula text of the current row cell at column `i`, or null. */
    getFormula(i: number): string | null {
        return this._currentFormulas[i] ?? null;
    }

    /** Rich-text runs of the current row cell at column `i`, or null. */
    getRichText(i: number): RichText | null {
        const inline = this._currentRichTexts[i];
        if (inline) return inline;
        const shared = this._currentSharedIndices[i];
        if (shared === undefined || shared < 0) return null;
        return this._sharedRichTexts[shared] ?? null;
    }

    private _resetCurrentRow(): void {
        this._currentRow = [];
        this._currentFormulas = [];
        this._currentStyleIndices = [];
        this._currentSharedIndices = [];
        this._currentRichTexts = [];
    }

    getValue(i: number): CellValue {
        if (i < 0 || i >= this._currentRow.length) return null;
        return this._currentRow[i];
    }
}

export default XlsxReader;
