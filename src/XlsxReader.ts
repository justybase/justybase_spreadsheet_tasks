import yauzl, { ZipFile, Entry } from 'yauzl';
import { ExcelReaderAbstract } from './ExcelReaderAbstract';
import { CellValue } from './Formats';
import { unescapeXml, columnLetterToIndex, parseSharedStringsXml, parseTagAttributes, resolveRelTarget } from './xmlUtils';
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

export class XlsxReader extends ExcelReaderAbstract {
    private zipfile: ZipFile | null = null;
    private sharedStrings: string[] = [];
    private sheetNames: string[] = [];
    private sheets: SheetInfo[] = [];
    private entries: Map<string, Entry> = new Map();

    private _currentSheetIndex: number = -1;
    private _sheetXml: string | null = null;
    private _xmlPos: number = 0;
    /** End of the sheetData element; row tags must not be searched past it. */
    private _xmlEnd: number = 0;

    private _currentRow: CellValue[] = [];
    private _expectedRowNum: number = 1;
    private _pendingEmptyCount: number = 0;

    xfIdToNumFmtId: number[] = [];
    customDateFormats: Set<number> = new Set();

    constructor() {
        super();
    }

    async open(path: string, readSharedStrings: boolean = true): Promise<void> {
        await this.close();
        // Reset state so re-opening does not accumulate sheets (xlspy parity).
        this.sheetNames = [];
        this.sheets = [];
        this.sharedStrings = [];
        this.xfIdToNumFmtId = [];
        this.customDateFormats = new Set();
        this.fieldCount = 0;
        this.resultsCount = 0;
        this.actualSheetName = '';
        this._currentSheetIndex = -1;
        this._sheetXml = null;
        this._xmlPos = 0;
        this._xmlEnd = 0;
        this._currentRow = [];
        this._expectedRowNum = 1;
        this._pendingEmptyCount = 0;
        this._oaEpoch = Date.UTC(1899, 11, 30);
        return new Promise((resolve, reject) => {
            yauzl.open(path, { lazyEntries: true, autoClose: false }, async (err, zipfile) => {
                if (err) return reject(err);
                this.zipfile = zipfile;

                this.entries = new Map();

                zipfile.on('entry', (entry: Entry) => {
                    this.entries.set(entry.fileName, entry);
                    zipfile.readEntry();
                });

                zipfile.on('end', async () => {
                    try {
                        const wbRelsContent = await this._readZipEntryContent('xl/_rels/workbook.xml.rels');
                        const rIdToTarget: Record<string, string> = {};
                        if (wbRelsContent) {
                            const relTagRe = /<Relationship\b[^>]*>/g;
                            let tagMatch;
                            while ((tagMatch = relTagRe.exec(wbRelsContent)) !== null) {
                                const attrs = parseTagAttributes(tagMatch[0]);
                                const id = attrs['Id'] ?? attrs['id'];
                                const target = attrs['Target'] ?? attrs['target'];
                                if (id !== undefined && target !== undefined) {
                                    rIdToTarget[id] = target;
                                }
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
                                const fullPath = resolveRelTarget(target);

                                this.sheetNames.push(name);
                                this.sheets.push({ name, sheetId, rId, path: fullPath });
                            }
                            // 1904 date system (workbookPr date1904="1").
                            const wbPrMatch = /<workbookPr\b[^>]*>/i.exec(wbContent);
                            if (wbPrMatch) {
                                const wbPrAttrs = parseTagAttributes(wbPrMatch[0]);
                                const v = (wbPrAttrs['date1904'] ?? '').trim().toLowerCase();
                                this._oaEpoch = (v === '1' || v === 'true') ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
                            }
                        }

                        if (readSharedStrings) {
                            const ssContent = await this._readZipEntryContent('xl/sharedStrings.xml');
                            if (ssContent) {
                                this._parseSharedStrings(ssContent);
                            }
                        }

                        const stylesContent = await this._readZipEntryContent('xl/styles.xml');
                        if (stylesContent) {
                            this._parseStyles(stylesContent);
                        }

                        this.resultsCount = this.sheets.length;
                        this._currentSheetIndex = -1;
                        resolve();
                    } catch (e) {
                        reject(e);
                    }
                });

                zipfile.readEntry();
            });
        });
    }

    async close(): Promise<void> {
        if (this.zipfile) {
            this.zipfile.close();
        }
    }

    private async _readZipEntryContent(path: string): Promise<string | null> {
        const entry = this.entries.get(path);
        if (!entry) return null;

        return new Promise((resolve, reject) => {
            this.zipfile!.openReadStream(entry, (err, readStream) => {
                if (err) return reject(err);

                const chunks: Buffer[] = [];
                readStream!.on('data', (chunk: Buffer) => chunks.push(chunk));
                readStream!.on('end', () => {
                    resolve(Buffer.concat(chunks).toString('utf8'));
                });
                readStream!.on('error', reject);
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
            const code = match[2];
            if (code.toLowerCase().includes('yy') || code.toLowerCase().includes('mm') || code.toLowerCase().includes('dd') || code.toLowerCase().includes('h:mm')) {
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
                    const numFmtId = parseInt(match[1]);
                    this.xfIdToNumFmtId.push(numFmtId);
                }
            }
        }
    }

    getSheetNames(): string[] {
        return this.sheetNames;
    }

    async read(): Promise<boolean> {
        if (this._currentSheetIndex === -1) {
            this._currentSheetIndex = 0;
            await this._initSheet(this._currentSheetIndex);
        }

        if (this._readNextRow()) {
            return true;
        } else {
            return false;
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

        this._sheetXml = await this._readZipEntryContent(sheet.path);

        if (!this._sheetXml) {
            this._sheetXml = "";
        }

        this._xmlPos = 0;
        this._xmlEnd = this._sheetXml.length;
        this._expectedRowNum = 1;
        this._pendingEmptyCount = 0;
        this._currentRow = [];
        this.fieldCount = 0;
        const sd = this._sheetXml.indexOf('<sheetData>');
        if (sd !== -1) {
            this._xmlPos = sd + 11;
            const se = this._sheetXml.indexOf('</sheetData>', this._xmlPos);
            if (se !== -1) this._xmlEnd = se;
        } else {
            // Namespaced or self-closed sheetData.
            const m = /<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/i.exec(this._sheetXml);
            if (m && m.index !== undefined) {
                this._xmlPos = m.index + m[0].length;
                if (/\/>\s*$/.test(m[0])) {
                    // <sheetData/> has no rows; do not scan the rest of the part.
                    this._xmlEnd = this._xmlPos;
                } else {
                    const se = this._sheetXml.indexOf('</sheetData>', this._xmlPos);
                    if (se !== -1) this._xmlEnd = se;
                }
            }
        }
    }

    private _readNextRow(): boolean {
        if (!this._sheetXml) return false;

        // Emit queued gap rows first (leading + intermediate). They are empty.
        if (this._pendingEmptyCount > 0) {
            this._pendingEmptyCount--;
            this._currentRow = [];
            this._expectedRowNum++;
            return true;
        }

        const rowStart = this._findRowTagStart(this._xmlPos);
        if (rowStart === -1) return false;

        const rowTagEnd = this._sheetXml.indexOf('>', rowStart);
        if (rowTagEnd === -1) return false;
        const rowTag = this._sheetXml.substring(rowStart, rowTagEnd + 1);
        const isSelfClosing = this._sheetXml.charCodeAt(rowTagEnd - 1) === 47; // '/>'
        const rowAttrs = parseTagAttributes(rowTag);
        const rRaw = rowAttrs['r'];
        let rowNum = this._expectedRowNum;
        if (rRaw !== undefined) {
            const parsed = parseInt(rRaw, 10);
            // Values beyond Excel's row limit are malformed; ignore them rather than
            // expanding the gap into millions of phantom rows.
            if (Number.isFinite(parsed) && parsed >= 1 && parsed <= MAX_ROWS) rowNum = parsed;
        }

        // Gap: rows without <row> elements are returned as empty rows (xlspy parity).
        if (rowNum > this._expectedRowNum) {
            const gap = rowNum - this._expectedRowNum;
            // Return one empty now, queue the rest without consuming the row.
            this._pendingEmptyCount = gap - 1;
            this._currentRow = [];
            this._expectedRowNum++;
            return true;
        }
        if (rowNum < this._expectedRowNum) {
            // Out-of-order or duplicate: consume and resync forward.
            if (isSelfClosing) {
                this._xmlPos = rowTagEnd + 1;
                this._currentRow = [];
                return true;
            }
            const rowEnd = this._sheetXml.indexOf('</row>', rowStart);
            if (rowEnd === -1) return false;
            this._xmlPos = rowEnd + 6;
            this._parseRowByIndex(rowStart, rowEnd);
            this._expectedRowNum = Math.max(this._expectedRowNum + 1, rowNum + 1);
            return true;
        }

        if (isSelfClosing) {
            this._xmlPos = rowTagEnd + 1;
            this._currentRow = [];
            this._expectedRowNum = rowNum + 1;
            return true;
        }

        const rowEnd = this._sheetXml.indexOf('</row>', rowStart);
        if (rowEnd === -1) return false;

        this._xmlPos = rowEnd + 6;
        this._parseRowByIndex(rowStart, rowEnd);
        this._expectedRowNum = rowNum + 1;
        return true;
    }

    /**
     * Find the next `<row …>` start, ignoring similarly named elements such as
     * `<rowBreaks>` that follow `</sheetData>`.
     */
    private _findRowTagStart(from: number): number {
        const xml = this._sheetXml!;
        let pos = from;
        while (pos < this._xmlEnd) {
            const idx = xml.indexOf('<row', pos);
            if (idx === -1 || idx >= this._xmlEnd) return -1;
            const next = xml.charCodeAt(idx + 4);
            // Whitespace, '/' or '>' means this is the row element, not e.g. <rowBreaks>.
            if (next === 0x3e || next === 0x2f || next === 0x20 ||
                next === 0x09 || next === 0x0a || next === 0x0d) {
                return idx;
            }
            pos = idx + 4;
        }
        return -1;
    }

    private _parseRowByIndex(rowStart: number, rowEnd: number): void {
        this._currentRow = [];
        const xml = this._sheetXml!;

        let pos = xml.indexOf('>', rowStart) + 1;
        let nextCol = 0;

        while (pos < rowEnd) {
            const cStart = xml.indexOf('<c', pos);
            if (cStart === -1 || cStart >= rowEnd) break;

            const nextTagClose = xml.indexOf('>', cStart);

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
            if (xml.charCodeAt(nextTagClose - 1) === 47) {
                pos = nextTagClose + 1;
            } else {
                const cEnd = xml.indexOf('</c>', nextTagClose);
                const cellContentEnd = (cEnd !== -1 && cEnd < rowEnd) ? cEnd : rowEnd;

                const vStart = xml.indexOf('<v>', nextTagClose);
                if (vStart !== -1 && vStart < cellContentEnd) {
                    const vInnerStart = vStart + 3;
                    const vEnd = xml.indexOf('</v>', vInnerStart);
                    if (vEnd !== -1 && vEnd < cellContentEnd) {
                        val = xml.substring(vInnerStart, vEnd);
                    }
                } else {
                    if (type === 'inlineStr') {
                        // Rich-text inlineStr: concatenate every <t> inside <is> (xlspy parity).
                        inlineVals = [];
                        let tSearch = nextTagClose;
                        while (true) {
                            const tStart = xml.indexOf('<t', tSearch);
                            if (tStart === -1 || tStart >= cellContentEnd) break;
                            const tClose = xml.indexOf('>', tStart);
                            if (tClose === -1 || tClose >= cellContentEnd) break;
                            const tContentStart = tClose + 1;
                            const tEnd = xml.indexOf('</t>', tContentStart);
                            if (tEnd === -1 || tEnd > cellContentEnd) break;
                            inlineVals.push(xml.substring(tContentStart, tEnd));
                            tSearch = tEnd + 4;
                        }
                        if (inlineVals.length === 0) inlineVals = null;
                    }
                }
                pos = cellContentEnd + 4;
            }

            let finalVal: any = null;
            if (inlineVals !== null) {
                finalVal = unescapeXml(inlineVals.join(''));
            } else if (val !== null) {
                if (type === 's') {
                    const idx = parseInt(val, 10);
                    finalVal = this.sharedStrings[idx];
                } else if (type === 'b') {
                    finalVal = (val === '1' || val === 'true');
                } else if (type === 'inlineStr') {
                    finalVal = unescapeXml(val);
                } else if (type === 'str') {
                    finalVal = unescapeXml(val);
                } else if (type === 'e') {
                    finalVal = unescapeXml(val);
                } else if (type === 'd') {
                    finalVal = parseIsoDateUtc(val) ?? unescapeXml(val);
                } else if (type === 'n' || type === '') {
                    finalVal = parseFloat(val);

                    if (styleIndex > 0 && this.xfIdToNumFmtId) {
                        let numFmtId = 0;
                        if (styleIndex < this.xfIdToNumFmtId.length) {
                            numFmtId = this.xfIdToNumFmtId[styleIndex];
                        }

                        const isDate = (numFmtId >= 14 && numFmtId <= 22) ||
                            (numFmtId >= 45 && numFmtId <= 47) ||
                            (this.customDateFormats && this.customDateFormats.has(numFmtId));

                        if (isDate) {
                            try {
                                finalVal = this.getDateTimeFromOaDate(finalVal);
                            } catch {
                                // ignore date conversion errors
                            }
                        }
                    }
                }
            }

            // The r attribute is optional: a cell without it sits in the column after the previous one.
            if (colIndex === -1) colIndex = nextCol;
            nextCol = colIndex + 1;

            this._currentRow[colIndex] = finalVal;
            if (colIndex >= this.fieldCount) this.fieldCount = colIndex + 1;
        }
    }

    getValue(i: number): CellValue {
        if (i < 0 || i >= this._currentRow.length) return null;
        return this._currentRow[i];
    }
}

export default XlsxReader;
