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
    cellRefAttr,
    assertCompressionLevel,
    DEFAULT_COMPRESSION_LEVEL,
} from './writerHelpers';
import { columnIndexToLetter } from './xmlUtils';

interface SheetInfo {
    name: string;
    pathInArchive: string;
    hidden: boolean;
    nameInArchive: string;
    sheetId: number;
    rId: string;
    filterHeaderRange: string | null;
}

/**
 * XLSX (Office Open XML) writer with streaming ZIP output.
 */
export class XlsxWriter {
    private output: fs.WriteStream;
    private archive: archiver.Archiver;
    private readonly _targetPath: string;
    private readonly _tempPath: string;
    private _outputError: Error | null = null;
    private _finalized: boolean = false;

    private sheetCount: number = 0;
    private sheetList: SheetInfo[] = [];
    private sstArray: string[] = [];
    private sstMap: Map<string, number> = new Map();
    private sstCntAll: number = 0;
    private _autofilterIsOn: boolean = false;

    private _oaEpoch: number;

    // Format registry
    private _formatRegistry: Map<string, number> = new Map();
    private _formatXfMap: Map<string, number> = new Map();
    private _nextNumfmtId: number = 165;
    private _nextXfIndex: number = 4;

    private readonly stream = new StreamingSheetState();
    private currentColLetters: string[] = [];

    /**
     * @param filePath Destination `.xlsx` path on disk.
     */
    /**
     * @param filePath Destination `.xlsx` path on disk.
     * @param options.compressionLevel Deflate level 0-9 (default 6).
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
    }

    private _getColumnLetter(colIndex: number): string {
        return columnIndexToLetter(colIndex);
    }

    /** @deprecated Internal; kept for unit-test access. Prefer sanitizeSheetName(). */
    private _sanitizeSheetName(name: string): string {
        return sanitizeSheetName(name, this.sheetCount);
    }

    private _escapeSheetNameForXml(name: string): string {
        return name
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    private _formatSheetNameForFormula(name: string): string {
        if (/[\s\-+=()!@#$%^&]/.test(name) || /^[0-9]/.test(name)) {
            const escaped = name.replace(/'/g, "''");
            return `'${escaped}'`;
        }
        return name;
    }

    /**
     * Register a worksheet (metadata only). Prefer `writeSheet` / `startSheet` for data.
     */
    addSheet(sheetName: string, hidden: boolean = false): void {
        const sanitizedName = sanitizeSheetName(sheetName, this.sheetCount);

        this.sheetCount++;
        const rId = `rId${this.sheetCount}`;
        const sheetFileName = `sheet${this.sheetCount}.xml`;
        this.sheetList.push({
            name: sanitizedName,
            pathInArchive: `xl/worksheets/${sheetFileName}`,
            hidden: hidden,
            nameInArchive: sheetFileName,
            sheetId: this.sheetCount,
            rId: rId,
            filterHeaderRange: null
        });
    }

    private _needsEscape(str: string): boolean {
        for (let i = 0; i < str.length; i++) {
            const c = str.charCodeAt(i);
            if (c === 38 || c === 60 || c === 62 || c === 34 || c === 39 || c === 13 ||
                (c >= 0 && c <= 8) || c === 11 || c === 12 || (c >= 14 && c <= 31)) {
                return true;
            }
        }
        return false;
    }

    private _escape(str: string): string {
        if (typeof str !== 'string') return str;
        // Literal `_xHHHH_` is written as `_x005F_xHHHH_` so Excel does not decode it (ST_Xstring).
        str = str.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_');
        if (!this._needsEscape(str)) return str;

        return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;')
            // CR must be a character reference: raw CR is normalized to LF by XML parsers.
            .replace(/\r/g, '&#13;')
            // eslint-disable-next-line no-control-regex
            .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
    }

    /**
     * Start streaming a sheet. Call `writeRow` for each row, then `endSheet`.
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

        this.addSheet(sheetName, hidden);

        this.stream.begin(columnCount, doAutofilter && headers !== undefined, new BigBuffer());
        const bigBuf = this.stream.buffer!;

        const colLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) {
            colLetters[i] = this._getColumnLetter(i);
        }
        this.currentColLetters = colLetters;

        if (headers) {
            applyHeaderWidths(this.stream.colWidths, headers, columnCount);
        }
        if (sampleRows) {
            updateColWidthsFromRows(this.stream.colWidths, sampleRows);
        }

        bigBuf.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        bigBuf.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

        // Dimension - just A1 since we don't know row count in advance
        bigBuf.writeString('<dimension ref="A1"/>');

        const isFirstSheet = this.sheetCount === 1;
        if (headers) {
            if (isFirstSheet) {
                bigBuf.writeString('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            } else {
                bigBuf.writeString('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            }
        } else {
            if (isFirstSheet) {
                bigBuf.writeString('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            } else {
                bigBuf.writeString('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            }
        }

        bigBuf.writeString('<cols>');
        for (let i = 0; i < columnCount; i++) {
            const width = defaultColWidth(this.stream.colWidths[i]);
            bigBuf.writeString(`<col min="${i + 1}" max="${i + 1}" width="${width}" bestFit="1" customWidth="1" />`);
        }
        bigBuf.writeString('</cols><sheetData>');

        if (headers) {
            this.stream.rowNum++;
            bigBuf.writeString(`<row r="${this.stream.rowNum}">`);
            // Headers are contiguous from column A, so no cell needs `r`.
            for (let c = 0; c < headers.length; c++) {
                this._writeStringCell(bigBuf, headers[c], '', 3);
            }
            bigBuf.writeString('</row>');
        }
    }

    /** Write one data row in streaming mode. */
    writeRow(row: CellValue[]): void {
        const bigBuf = this.stream.assertStreaming();

        if (row.length !== this.stream.endCol - this.stream.startCol) {
            throw new Error(
                `Row length mismatch. Expected ${this.stream.endCol - this.stream.startCol} columns, got ${row.length}`
            );
        }

        this.stream.rowNum++;
        assertRowCount(this.stream.rowNum);
        bigBuf.writeString(`<row r="${this.stream.rowNum}">`);

        let prevCol = -1;
        for (let c = 0; c < row.length; c++) {
            const raw = row[c];
            if (raw === null || raw === undefined) continue;

            const fmtString = getFormat(raw);
            const val = fmtString !== null ? unwrapCell(raw) : raw;
            const xfIdx = fmtString !== null ? this._registerFormat(fmtString) : -1;
            const styleAttr = xfIdx >= 0 ? ` s="${xfIdx}"` : '';

            const ref = cellRefAttr(c, prevCol, this.currentColLetters[c], this.stream.rowNum);
            prevCol = c;
            if (typeof val === 'number') {
                if (Number.isFinite(val)) {
                    bigBuf.writeString(`<c${ref}${styleAttr}><v>${val}</v></c>`);
                } else {
                    this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                }
            } else if (typeof val === 'bigint') {
                this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
            } else if (typeof val === 'boolean') {
                bigBuf.writeString(`<c${ref} t="b"${styleAttr}><v>${val ? 1 : 0}</v></c>`);
            } else if (val instanceof Date) {
                const oaDate = this._toOADate(val);
                if (Number.isFinite(oaDate)) {
                    const finalStyle = xfIdx >= 0 ? xfIdx : 1;
                    bigBuf.writeString(`<c${ref} s="${finalStyle}"><v>${oaDate}</v></c>`);
                } else {
                    this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                }
            } else if (val !== null && val !== undefined) {
                this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
            }
        }
        bigBuf.writeString('</row>');
    }

    /** Finish the current streaming sheet and append it to the archive. */
    endSheet(): void {
        const bigBuf = this.stream.assertStreaming();
        bigBuf.writeString('</sheetData>');

        if (this.stream.doAutofilter && this.stream.endCol > 0) {
            this._autofilterIsOn = true;
            const filterRef = `A1:${this.currentColLetters[this.stream.endCol - 1]}${this.stream.rowNum}`;
            bigBuf.writeString(`<autoFilter ref="${filterRef}"/>`);

            const sheet = this.sheetList[this.sheetCount - 1];
            const formulaSheetName = this._formatSheetNameForFormula(sheet.name);
            sheet.filterHeaderRange = `${formulaSheetName}!$A$1:$${this.currentColLetters[this.stream.endCol - 1]}$${this.stream.rowNum}`;
        }

        bigBuf.writeString('</worksheet>');

        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: this.sheetList[this.sheetCount - 1].pathInArchive
        });

        this.stream.end();
        this.currentColLetters = [];
    }

    /**
     * Write an entire sheet in one call (batch mode).
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
        const totalRows = rows.length + (headers ? 1 : 0);
        assertRowCount(totalRows);
        if (headers) {
            assertRowFitsColumns('header', headers.length, columnCount);
            for (const h of headers) assertCellTextLength(h);
        }
        for (let r = 0; r < rows.length; r++) {
            assertRowFitsColumns('row', rows[r].length, columnCount);
        }

        const colWidths = initColWidths(columnCount);

        const colLetters = new Array(columnCount);
        for (let i = 0; i < columnCount; i++) {
            colLetters[i] = this._getColumnLetter(i);
        }

        if (headers) {
            applyHeaderWidths(colWidths, headers, columnCount);
        }
        updateColWidthsFromRows(colWidths, rows);

        bigBuf.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        bigBuf.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

        if (totalRows > 0 && columnCount > 0) {
            bigBuf.writeString(`<dimension ref="A1:${colLetters[columnCount - 1]}${totalRows}"/>`);
        } else {
            bigBuf.writeString('<dimension ref="A1"/>');
        }

        const isFirstSheet = this.sheetCount === 1;
        if (headers) {
            if (isFirstSheet) {
                bigBuf.writeString('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            } else {
                bigBuf.writeString('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            }
        } else {
            if (isFirstSheet) {
                bigBuf.writeString('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            } else {
                bigBuf.writeString('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
            }
        }

        bigBuf.writeString('<cols>');
        for (let i = 0; i < columnCount; i++) {
            const width = defaultColWidth(colWidths[i]);
            bigBuf.writeString(`<col min="${i + 1}" max="${i + 1}" width="${width}" bestFit="1" customWidth="1" />`);
        }
        bigBuf.writeString('</cols><sheetData>');

        let rowNum = 0;

        if (headers) {
            rowNum++;
            bigBuf.writeString(`<row r="${rowNum}">`);
            // Headers are contiguous from column A, so no cell needs `r`.
            for (let c = 0; c < headers.length; c++) {
                this._writeStringCell(bigBuf, headers[c], '', 3);
            }
            bigBuf.writeString('</row>');
        }

        for (let r = 0; r < rows.length; r++) {
            rowNum++;
            bigBuf.writeString(`<row r="${rowNum}">`);
            const row = rows[r];
            let prevCol = -1;
            for (let c = 0; c < row.length; c++) {
                const raw = row[c];
                if (raw === null || raw === undefined) continue;

                const fmtString = getFormat(raw);
                const val = fmtString !== null ? unwrapCell(raw) : raw;
                const xfIdx = fmtString !== null ? this._registerFormat(fmtString) : -1;
                const styleAttr = xfIdx >= 0 ? ` s="${xfIdx}"` : '';

                const ref = cellRefAttr(c, prevCol, colLetters[c], rowNum);
                prevCol = c;
                if (typeof val === 'number') {
                    if (Number.isFinite(val)) {
                        bigBuf.writeString(`<c${ref}${styleAttr}><v>${val}</v></c>`);
                    } else {
                        this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                    }
                } else if (typeof val === 'bigint') {
                    this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                } else if (typeof val === 'boolean') {
                    bigBuf.writeString(`<c${ref} t="b"${styleAttr}><v>${val ? 1 : 0}</v></c>`);
                } else if (val instanceof Date) {
                    const oaDate = this._toOADate(val);
                    if (Number.isFinite(oaDate)) {
                        const finalStyle = xfIdx >= 0 ? xfIdx : 1;
                        bigBuf.writeString(`<c${ref} s="${finalStyle}"><v>${oaDate}</v></c>`);
                    } else {
                        this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                    }
                } else if (val !== null && val !== undefined) {
                    this._writeStringCell(bigBuf, val.toString(), ref, xfIdx >= 0 ? xfIdx : undefined);
                }
            }
            bigBuf.writeString('</row>');
        }

        bigBuf.writeString('</sheetData>');

        if (doAutofilter && headers && columnCount > 0) {
            this._autofilterIsOn = true;
            const filterRef = `A1:${colLetters[columnCount - 1]}${totalRows}`;
            bigBuf.writeString(`<autoFilter ref="${filterRef}"/>`);

            const sheet = this.sheetList[this.sheetCount - 1];
            const formulaSheetName = this._formatSheetNameForFormula(sheet.name);
            sheet.filterHeaderRange = `${formulaSheetName}!$A$1:$${colLetters[columnCount - 1]}$${totalRows}`;
        }

        bigBuf.writeString('</worksheet>');

        this.archive.append(Readable.from(bigBuf.getChunks()), {
            name: this.sheetList[this.sheetCount - 1].pathInArchive
        });
    }

    /** `ref` is the cell's `r` attribute from `cellRefAttr`, or '' when omitted. */
    private _writeStringCell(bigBuf: BigBuffer, val: string, ref: string, styleOverride?: number): void {
        assertCellTextLength(val);
        let index = this.sstMap.get(val);
        if (index === undefined) {
            index = this.sstArray.length;
            this.sstArray.push(val);
            this.sstMap.set(val, index);
        }
        this.sstCntAll++;
        const styleAttr = styleOverride !== undefined ? ` s="${styleOverride}"` : '';
        bigBuf.writeString(`<c${ref} t="s"${styleAttr}><v>${index}</v></c>`);
    }

    private _registerFormat(fmtString: string): number {
        if (this._formatXfMap.has(fmtString)) {
            return this._formatXfMap.get(fmtString)!;
        }
        const numfmtId = this._nextNumfmtId++;
        this._formatRegistry.set(fmtString, numfmtId);
        const xfIndex = this._nextXfIndex++;
        this._formatXfMap.set(fmtString, xfIndex);
        return xfIndex;
    }

    private _toOADate(date: Date): number {
        // Use UTC time to avoid timezone conversion issues; handles the
        // fictitious 1900-02-29 (dates before 1900-03-01 serialize one lower).
        return timeToOaDate(date.getTime(), this._oaEpoch);
    }

    private _getDateWidth(_date: Date): number {
        return 19; // "YYYY-MM-DD HH:MM:SS" ≈ 19 chars
    }

    /** Finalize the ZIP package and close the output stream. */
    finalize(): Promise<void> {
        return new Promise((resolve, reject) => {
            if (this._finalized) {
                reject(new Error('XlsxWriter: already finalized'));
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
                    // Ensure data hits disk before publishing (xlspy parity, no fsync of dir).
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
                this._writeSharedStrings();
                this._writeStyles();
                this._writeWorkbook();
                this._writeContentTypes();
                this._writeRels();

                this.output.once('close', succeed);
                this.output.once('error', fail);
                this.archive.once('error', fail);

                this.archive.finalize();
            } catch (err) {
                fail(err);
            }
        });
    }

    private _writeSharedStrings(): void {
        const bigBuf = new BigBuffer();

        bigBuf.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        bigBuf.writeString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.sstCntAll}" uniqueCount="${this.sstArray.length}">`);

        for (let i = 0; i < this.sstArray.length; i++) {
            const txt = this.sstArray[i];
            const cleanTxt = this._escape(txt);
            if (cleanTxt.length > 0 && (cleanTxt[0] === ' ' || cleanTxt[cleanTxt.length - 1] === ' ' || /[\t\n\r]/.test(cleanTxt))) {
                bigBuf.writeString(`<si><t xml:space="preserve">${cleanTxt}</t></si>`);
            } else {
                bigBuf.writeString(`<si><t>${cleanTxt}</t></si>`);
            }
        }

        bigBuf.writeString('</sst>');

        this.archive.append(Readable.from(bigBuf.getChunks()), { name: 'xl/sharedStrings.xml' });
    }

    private _writeStyles(): void {
        let numFmts = '';
        let numFmtCount = 0;

        const xfEntries: string[] = [];
        xfEntries.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
        xfEntries.push('<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>');
        xfEntries.push('<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>');
        xfEntries.push('<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>');

        for (const [fmtString, numfmtId] of this._formatRegistry) {
            const escapedFmt = fmtString
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
            numFmts += `<numFmt numFmtId="${numfmtId}" formatCode="${escapedFmt}"/>`;
            numFmtCount++;
        }

        const sortedFormats = [...this._formatXfMap.entries()].sort((a, b) => a[1] - b[1]);
        for (const [fmtString] of sortedFormats) {
            const numfmtId = this._formatRegistry.get(fmtString)!;
            xfEntries.push(`<xf numFmtId="${numfmtId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`);
        }

        const totalXf = xfEntries.length;
        const numFmtsSection = numFmtCount > 0 ? `<numFmts count="${numFmtCount}">${numFmts}</numFmts>` : '';

        const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${numFmtsSection}
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="${totalXf}">
${xfEntries.join('\n')}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
        this.archive.append(styles, { name: 'xl/styles.xml' });
    }

    private _writeWorkbook(): void {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>`;

        for (const sheet of this.sheetList) {
            const state = sheet.hidden ? ' state="hidden"' : '';
            const escapedName = this._escapeSheetNameForXml(sheet.name);
            xml += `<sheet name="${escapedName}" sheetId="${sheet.sheetId}"${state} r:id="${sheet.rId}"/>`;
        }

        xml += `</sheets>`;

        if (this._autofilterIsOn) {
            xml += '<definedNames>';
            for (const sheet of this.sheetList) {
                if (sheet.filterHeaderRange) {
                    const localSheetId = sheet.sheetId - 1;
                    const escapedRange = this._escape(sheet.filterHeaderRange);
                    xml += `<definedName name="_xlnm._FilterDatabase" localSheetId="${localSheetId}" hidden="1">${escapedRange}</definedName>`;
                }
            }
            xml += '</definedNames>';
        }

        xml += '<calcPr calcId="124519" fullCalcOnLoad="1"/>';
        xml += `</workbook>`;
        this.archive.append(xml, { name: 'xl/workbook.xml' });
    }

    private _writeContentTypes(): void {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;

        for (const sheet of this.sheetList) {
            xml += `<Override PartName="/${sheet.pathInArchive}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        }

        xml += `</Types>`;
        this.archive.append(xml, { name: '[Content_Types].xml' });
    }

    private _writeRels(): void {
        const globalRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
        this.archive.append(globalRels, { name: '_rels/.rels' });

        let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

        for (const sheet of this.sheetList) {
            wbRels += `<Relationship Id="${sheet.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheet.nameInArchive}"/>`;
        }

        let nextId = this.sheetList.length + 1;
        wbRels += `<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
        wbRels += `<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;

        wbRels += `</Relationships>`;
        this.archive.append(wbRels, { name: 'xl/_rels/workbook.xml.rels' });
    }
}

export default XlsxWriter;
