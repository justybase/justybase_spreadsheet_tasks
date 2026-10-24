/**
 * BrowserXlsxWriter - Browser XLSX writer with bold/frozen headers and autofit.
 * @module BrowserXlsxWriter
 */
import { BrowserBigBuffer } from './BrowserBigBuffer';
import { BrowserZip } from './BrowserZip';
import { CellValue, isFormattedCell, unwrapCell, getFormat } from '../Formats';
import { columnIndexToLetter } from '../xmlUtils';
import {
    assertCellTextLength as assertBrowserCellText,
    assertColumnCount as assertBrowserColumnCount,
    assertRowFitsColumns as assertBrowserRowFits,
    assertRowCount as assertBrowserRowCount,
    timeToOaDate as timeToOaDateBrowser,
    cellRefAttr,
} from '../writerHelpers';

export type BrowserHeaderStyle = 'bold' | 'fill' | 'bold+fill';

export interface BrowserSheetOptions {
    doAutofilter?: boolean;
    headerStyle?: BrowserHeaderStyle;
}

export interface BrowserStartSheetOptions extends BrowserSheetOptions {
    hidden?: boolean;
}

interface BrowserSheetEntry {
    name: string;
    path: string;
    hidden: boolean;
    fn: string;
    id: number;
    rId: string;
    fhr: string | null;
}

export class BrowserXlsxWriter {
    private _zip: BrowserZip;
    private _sc: number;
    private _sl: BrowserSheetEntry[];
    private _sstA: string[];
    private _sstM: Map<string, number>;
    private _sstCnt: number;
    private _cw: number[];
    private _afOn: boolean;
    private _oaE: number;
    private _csb: BrowserBigBuffer | null;
    private _csrn: number;
    private _csSC: number;
    private _csEC: number;
    private _csAF: boolean;
    private _isS: boolean;
    private _csCL: string[];
    private _fmtReg: Map<string, number>;
    private _fmtXf: Map<string, number>;
    private _nextNfmt: number;
    private _nextXf: number;

    constructor() {
        this._zip = new BrowserZip();
        this._sc = 0; this._sl = []; this._sstA = []; this._sstM = new Map();
        this._sstCnt = 0; this._cw = []; this._afOn = false;
        this._oaE = Date.UTC(1899, 11, 30);
        this._csb = null; this._csrn = 0; this._csSC = 0; this._csEC = 0;
        this._csAF = false; this._isS = false; this._csCL = [];
        this._fmtReg = new Map();
        this._fmtXf = new Map();
        this._nextNfmt = 165;
        this._nextXf = 6;
    }
    _cl(i: number): string { return columnIndexToLetter(i); }
    _san(n: string): string {
        if (!n || typeof n !== 'string') return `Sheet${this._sc + 1}`;
        let s = n.replace(/[\\/*?[\]:]/g, '_'); if (s.length > 31) s = s.substring(0, 31);
        if (!s.trim()) s = `Sheet${this._sc + 1}`; return s;
    }
    _ex(s: string): string { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;'); }
    _esc(s: string): string {
        // Literal `_xHHHH_` is written as `_x005F_xHHHH_` so Excel does not decode it (ST_Xstring).
        s = s.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_');
        let need = false;
        for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); if (c === 38 || c === 60 || c === 62 || c === 34 || c === 39 || c === 13 || (c >= 0 && c <= 8) || c === 11 || c === 12 || (c >= 14 && c <= 31)) { need = true; break; } }
        if (!need) return s;
        // CR is a character reference: raw CR is normalized to LF by XML parsers.
        const escaped = s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;').replace(/\r/g, '&#13;');
        let out = '';
        for (let i = 0; i < escaped.length; i++) {
            const c = escaped.charCodeAt(i);
            if (c <= 8 || c === 11 || c === 12 || (c >= 14 && c <= 31)) continue;
            out += escaped[i];
        }
        return out;
    }
    _fmn(n: string): string { return /[\s\-+=()!@#$%^&]/.test(n) || /^[0-9]/.test(n) ? `'${n.replace(/'/g, "''")}'` : n; }

    addSheet(name: string, hidden = false): void {
        const sn = this._san(name); this._sc++;
        this._sl.push({ name: sn, path: `xl/worksheets/sheet${this._sc}.xml`, hidden, fn: `sheet${this._sc}.xml`, id: this._sc, rId: `rId${this._sc}`, fhr: null });
    }
    /** `ref` is the cell's `r` attribute from `cellRefAttr`, or '' when omitted. */
    _wsc(bb: BrowserBigBuffer, val: string, ref: string, so?: number): void {
        assertBrowserCellText(val);
        let idx = this._sstM.get(val);
        if (idx === undefined) { idx = this._sstA.length; this._sstA.push(val); this._sstM.set(val, idx); }
        this._sstCnt++;
        const sa = so !== undefined ? ` s="${so}"` : '';
        bb.writeString(`<c${ref} t="s"${sa}><v>${idx}</v></c>`);
    }
    _wscStyle(bb: BrowserBigBuffer, val: string, ref: string, styleId: number): void {
        assertBrowserCellText(val);
        let idx = this._sstM.get(val);
        if (idx === undefined) { idx = this._sstA.length; this._sstA.push(val); this._sstM.set(val, idx); }
        this._sstCnt++; bb.writeString(`<c${ref} t="s" s="${styleId}"><v>${idx}</v></c>`);
    }
    _oa(d: Date): number { return timeToOaDateBrowser(d.getTime(), this._oaE); }
    _regFmt(fs: string): number {
        const existing = this._fmtXf.get(fs);
        if (existing !== undefined) return existing;
        const nid = this._nextNfmt++;
        this._fmtReg.set(fs, nid);
        const xf = this._nextXf++;
        this._fmtXf.set(fs, xf);
        return xf;
    }
    /**
     * Write one cell of row `rn` at column `c` (letter `cr`). Returns the column
     * of the last cell written, which the caller passes as `prevCol` for the next
     * cell; a null cell writes nothing and leaves `prevCol` unchanged.
     */
    _wcv(bb: BrowserBigBuffer, raw: CellValue, c: number, cr: string, rn: number, prevCol: number): number {
        if (raw === null || raw === undefined) return prevCol;
        const fs = getFormat(raw);
        const v = unwrapCell(raw);
        const xf = fs !== null ? this._regFmt(fs) : -1;
        const sa = xf >= 0 ? ` s="${xf}"` : '';
        // A formatted cell may wrap null/undefined; Node XlsxWriter skips it the same way.
        if (v === null || v === undefined) return prevCol;
        const ref = cellRefAttr(c, prevCol, cr, rn);
        const so = xf >= 0 ? xf : void 0;
        if (typeof v === 'number') { if (Number.isFinite(v)) bb.writeString(`<c${ref}${sa}><v>${v}</v></c>`); else this._wsc(bb, v.toString(), ref, so); }
        else if (typeof v === 'bigint') this._wsc(bb, v.toString(), ref, so);
        else if (typeof v === 'boolean') bb.writeString(`<c${ref} t="b"${sa}><v>${v ? 1 : 0}</v></c>`);
        else if (v instanceof Date) { const o = this._oa(v); if (Number.isFinite(o)) bb.writeString(`<c${ref} s="${xf >= 0 ? xf : 1}"><v>${o}</v></c>`); else this._wsc(bb, v.toString(), ref, so); }
        else this._wsc(bb, v.toString(), ref, so);
        return c;
    }
    _sheetHead(bb: BrowserBigBuffer, cc: number, isFirst: boolean, hasAF: boolean): void {
        bb.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        bb.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
        bb.writeString('<dimension ref="A1"/>');
        if (hasAF) bb.writeString(`<sheetViews><sheetView ${isFirst ? 'tabSelected="1" ' : ''}workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>`);
        else bb.writeString(`<sheetViews><sheetView ${isFirst ? 'tabSelected="1" ' : ''}workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>`);
        bb.writeString('<cols>');
        for (let i = 0; i < cc; i++) { const w = this._cw[i] > 0 ? this._cw[i] : 10; bb.writeString(`<col min="${i + 1}" max="${i + 1}" width="${w}" bestFit="1" customWidth="1"/>`); }
        bb.writeString('</cols><sheetData>');
    }

    startSheet(name: string, colCount: number, headers?: string[], options: BrowserStartSheetOptions = {}): void {
        if (this._isS) throw new Error('Already streaming');
        assertBrowserColumnCount(colCount);
        if (headers !== undefined) { assertBrowserRowFits('header', headers.length, colCount); for (const h of headers) assertBrowserCellText(h); }
        const { hidden = false, doAutofilter = true, headerStyle = 'bold' } = options;
        let hs = 3; if (headerStyle === 'fill') hs = 4; else if (headerStyle === 'bold+fill') hs = 5;
        this.addSheet(name, hidden); this._isS = true;
        const csb = new BrowserBigBuffer();
        this._csb = csb; this._csrn = 0; this._csSC = 0; this._csEC = colCount;
        this._csAF = doAutofilter && headers !== undefined;
        this._cw = new Array(colCount).fill(-1);
        const cl: string[] = new Array(colCount); for (let i = 0; i < colCount; i++) cl[i] = this._cl(i); this._csCL = cl;
        if (headers) for (let i = 0; i < colCount; i++) { let w = 1.3 * (headers[i] ? headers[i].length : 0) + 3; if (w > 80) w = 80; if (this._cw[i] < w) this._cw[i] = w; }
        this._sheetHead(csb, colCount, this._sc === 1, this._csAF);
        if (headers) { this._csrn++; csb.writeString(`<row r="${this._csrn}">`); for (let c = 0; c < headers.length; c++) this._wscStyle(csb, headers[c], '', hs); csb.writeString('</row>'); }
    }
    writeRow(row: CellValue[]): void {
        if (!this._isS || this._csb === null) throw new Error('Not streaming');
        if (row.length !== this._csEC) throw new Error(`Row length mismatch. Expected ${this._csEC} columns, got ${row.length}`);
        const csb = this._csb;
        this._csrn++; assertBrowserRowCount(this._csrn); csb.writeString(`<row r="${this._csrn}">`);
        let prev = -1; for (let c = 0; c < row.length; c++) prev = this._wcv(csb, row[c], c, this._csCL[c], this._csrn, prev);
        csb.writeString('</row>');
    }
    endSheet(): void {
        if (!this._isS || this._csb === null) throw new Error('Not streaming');
        const csb = this._csb;
        csb.writeString('</sheetData>');
        if (this._csAF && this._csEC > 0) { this._afOn = true; const fr = `A1:${this._csCL[this._csEC - 1]}${this._csrn}`; csb.writeString(`<autoFilter ref="${fr}"/>`); const sh = this._sl[this._sc - 1]; sh.fhr = `${this._fmn(sh.name)}!$A$1:$${this._csCL[this._csEC - 1]}$${this._csrn}`; }
        csb.writeString('</worksheet>');
        this._zip.addFile(this._sl[this._sc - 1].path, csb.toUint8Array());
        this._isS = false; this._csb = null;
    }

    writeSheet(rows: CellValue[][], headers: string[] | null = null, options: BrowserSheetOptions = {}): void {
        const doAutofilter = options.doAutofilter !== false;
        const headerStyle = options.headerStyle || 'bold';
        let hs = 3; if (headerStyle === 'fill') hs = 4; else if (headerStyle === 'bold+fill') hs = 5;
        const bb = new BrowserBigBuffer();
        const cc = rows.length > 0 ? rows[0].length : (headers ? headers.length : 0);
        if (cc > 0) assertBrowserColumnCount(cc);
        const tr = rows.length + (headers ? 1 : 0); assertBrowserRowCount(tr);
        if (headers) { assertBrowserRowFits('header', headers.length, cc); for (const h of headers) assertBrowserCellText(h); }
        for (let r = 0; r < rows.length; r++) assertBrowserRowFits('row', rows[r].length, cc);
        this._cw = new Array(cc).fill(-1);
        const cl: string[] = new Array(cc); for (let i = 0; i < cc; i++) cl[i] = this._cl(i);
        if (headers) for (let i = 0; i < cc; i++) { let w = 1.3 * (headers[i] ? headers[i].length : 0) + 3; if (w > 80) w = 80; if (this._cw[i] < w) this._cw[i] = w; }
        for (let r = 0; r < Math.min(rows.length, 100); r++) for (let c = 0; c < rows[r].length; c++) {
            const raw = rows[r][c]; if (raw == null) continue;
            const v = isFormattedCell(raw) ? raw.value : raw;
            if (v === null || v === undefined) continue;
            let w = 1.3 * (v instanceof Date ? 10 : v.toString().length) + 3; if (w > 80) w = 80; if (this._cw[c] < w) this._cw[c] = w;
        }
        this._sheetHead(bb, cc, this._sc === 1, doAutofilter && !!headers);
        let rn = 0;
        if (headers) { rn++; bb.writeString(`<row r="${rn}">`); for (let c = 0; c < headers.length; c++) this._wscStyle(bb, headers[c], '', hs); bb.writeString('</row>'); }
        for (let r = 0; r < rows.length; r++) { rn++; bb.writeString(`<row r="${rn}">`); let prev = -1; for (let c = 0; c < rows[r].length; c++) prev = this._wcv(bb, rows[r][c], c, cl[c], rn, prev); bb.writeString('</row>'); }
        bb.writeString('</sheetData>');
        if (doAutofilter && headers && cc > 0) { this._afOn = true; const fr = `A1:${cl[cc - 1]}${tr}`; bb.writeString(`<autoFilter ref="${fr}"/>`); const sh = this._sl[this._sc - 1]; sh.fhr = `${this._fmn(sh.name)}!$A$1:$${cl[cc - 1]}$${tr}`; }
        bb.writeString('</worksheet>');
        this._zip.addFile(this._sl[this._sc - 1].path, bb.toUint8Array());
    }

    finalize(): Blob {
        this._writeSst(); this._writeStyles(); this._writeWb(); this._writeCt(); this._writeRels();
        return this._zip.toBlob();
    }
    _writeSst(): void {
        const bb = new BrowserBigBuffer();
        bb.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        bb.writeString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this._sstCnt}" uniqueCount="${this._sstA.length}">`);
        for (const t of this._sstA) { const c = this._esc(t); if (c.length > 0 && (c[0] === ' ' || c[c.length - 1] === ' ' || /[\t\n\r]/.test(c))) bb.writeString(`<si><t xml:space="preserve">${c}</t></si>`); else bb.writeString(`<si><t>${c}</t></si>`); }
        bb.writeString('</sst>'); this._zip.addFile('xl/sharedStrings.xml', bb.toUint8Array());
    }
    _writeStyles(): void {
        let nfs = '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/>';
        let nfc = 1;
        const xfe = [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            '<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>',
            '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        ];
        for (const [fs, nid] of this._fmtReg) {
            const ef = fs.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
            nfs += `<numFmt numFmtId="${nid}" formatCode="${ef}"/>`;
            nfc++;
        }
        const sf = [...this._fmtXf.entries()].sort((a, b) => a[1] - b[1]);
        for (const [fs] of sf) {
            const nid = this._fmtReg.get(fs);
            xfe.push(`<xf numFmtId="${nid}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`);
        }
        this._zip.addFile('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="${nfc}">${nfs}</numFmts>
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${xfe.length}">
${xfe.join('\n')}
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`);
    }
    _writeWb(): void {
        let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews><sheets>`;
        for (const s of this._sl) x += `<sheet name="${this._ex(s.name)}" sheetId="${s.id}"${s.hidden ? ' state="hidden"' : ''} r:id="${s.rId}"/>`;
        x += '</sheets>';
        if (this._afOn) { x += '<definedNames>'; for (const s of this._sl) if (s.fhr) x += `<definedName name="_xlnm._FilterDatabase" localSheetId="${s.id - 1}" hidden="1">${this._esc(s.fhr)}</definedName>`; x += '</definedNames>'; }
        x += '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>';
        this._zip.addFile('xl/workbook.xml', x);
    }
    _writeCt(): void {
        let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
        for (const s of this._sl) x += `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        x += '</Types>'; this._zip.addFile('[Content_Types].xml', x);
    }
    _writeRels(): void {
        this._zip.addFile('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
        let r = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
        for (const s of this._sl) r += `<Relationship Id="${s.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.fn}"/>`;
        let n = this._sl.length + 1;
        r += `<Relationship Id="rId${n++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
        r += `<Relationship Id="rId${n++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
        r += '</Relationships>'; this._zip.addFile('xl/_rels/workbook.xml.rels', r);
    }
}
export default BrowserXlsxWriter;
