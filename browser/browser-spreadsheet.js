// src/browser/BrowserBigBuffer.ts
var BrowserBigBuffer = class {
  /**
   * @param {number} [chunkSize=65536]
   */
  constructor(chunkSize = 65536) {
    this.chunkSize = chunkSize;
    this.chunks = [];
    this.currentBuffer = new Uint8Array(chunkSize);
    this.dataView = new DataView(this.currentBuffer.buffer);
    this.cursor = 0;
    this._textEncoder = new TextEncoder();
  }
  /**
   * @param {number} size
   */
  ensureCapacity(size) {
    if (this.cursor + size > this.chunkSize) {
      this._flush();
    }
  }
  _flush() {
    if (this.cursor > 0) {
      this.chunks.push(this.currentBuffer.slice(0, this.cursor));
      this.currentBuffer = new Uint8Array(this.chunkSize);
      this.dataView = new DataView(this.currentBuffer.buffer);
      this.cursor = 0;
    }
  }
  /**
   * Write a Uint8Array to the buffer.
   * @param {Uint8Array} buffer
   */
  write(buffer) {
    let len = buffer.length;
    let offset = 0;
    while (len > 0) {
      let available = this.chunkSize - this.cursor;
      if (available === 0) {
        this._flush();
        available = this.chunkSize;
      }
      const toWrite = Math.min(len, available);
      this.currentBuffer.set(buffer.subarray(offset, offset + toWrite), this.cursor);
      this.cursor += toWrite;
      offset += toWrite;
      len -= toWrite;
    }
  }
  /**
   * @param {number} val
   */
  writeByte(val) {
    this.ensureCapacity(1);
    this.currentBuffer[this.cursor] = val;
    this.cursor++;
  }
  /**
   * @param {number} val
   */
  writeUnsafeByte(val) {
    this.currentBuffer[this.cursor++] = val;
  }
  /**
   * @param {number} val
   */
  writeInt32LE(val) {
    this.ensureCapacity(4);
    this.dataView.setInt32(this.cursor, val, true);
    this.cursor += 4;
  }
  /**
   * @param {number} val
   */
  writeUnsafeInt32LE(val) {
    this.dataView.setInt32(this.cursor, val, true);
    this.cursor += 4;
  }
  /**
   * @param {number} val
   */
  writeDoubleLE(val) {
    this.ensureCapacity(8);
    this.dataView.setFloat64(this.cursor, val, true);
    this.cursor += 8;
  }
  /**
   * @param {number} val
   */
  writeUnsafeDoubleLE(val) {
    this.dataView.setFloat64(this.cursor, val, true);
    this.cursor += 8;
  }
  /**
   * Write a UTF-8 string.
   * @param {string} str
   */
  writeString(str) {
    const encoded = this._textEncoder.encode(str);
    const byteLength = encoded.length;
    if (this.cursor + byteLength <= this.chunkSize) {
      this.currentBuffer.set(encoded, this.cursor);
      this.cursor += byteLength;
      return;
    }
    this._flush();
    if (byteLength > this.chunkSize) {
      this.chunks.push(encoded);
    } else {
      this.currentBuffer.set(encoded, this.cursor);
      this.cursor += byteLength;
    }
  }
  /**
   * Write a string as UTF-16LE.
   * @param {string} str
   */
  writeUtf16LE(str) {
    const byteLength = str.length * 2;
    if (this.cursor + byteLength <= this.chunkSize) {
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        this.currentBuffer[this.cursor++] = code & 255;
        this.currentBuffer[this.cursor++] = code >> 8 & 255;
      }
      return;
    }
    this._flush();
    if (byteLength > this.chunkSize) {
      const buf = new Uint8Array(byteLength);
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        buf[i * 2] = code & 255;
        buf[i * 2 + 1] = code >> 8 & 255;
      }
      this.chunks.push(buf);
    } else {
      for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        this.currentBuffer[this.cursor++] = code & 255;
        this.currentBuffer[this.cursor++] = code >> 8 & 255;
      }
    }
  }
  /**
   * Flush and return all chunks.
   * @returns {Uint8Array[]}
   */
  getChunks() {
    if (this.cursor > 0) {
      this.chunks.push(this.currentBuffer.slice(0, this.cursor));
      this.currentBuffer = new Uint8Array(this.chunkSize);
      this.dataView = new DataView(this.currentBuffer.buffer);
      this.cursor = 0;
    }
    return this.chunks;
  }
  /**
   * Concatenate all chunks into a single Uint8Array.
   * @returns {Uint8Array}
   */
  toUint8Array() {
    const chunks = this.getChunks();
    let totalLen = 0;
    for (const c of chunks) totalLen += c.length;
    const result = new Uint8Array(totalLen);
    let offset = 0;
    for (const c of chunks) {
      result.set(c, offset);
      offset += c.length;
    }
    return result;
  }
  reset() {
    this.chunks = [];
    this.cursor = 0;
  }
};

// src/browser/BrowserZip.ts
var BrowserZip = class _BrowserZip {
  constructor() {
    this.files = [];
    this.files = [];
  }
  static {
    this._crc32Table = null;
  }
  /**
   * Add a file entry (string or Uint8Array).
   * @param {string} name - Path inside the ZIP (e.g. "xl/workbook.xml")
   * @param {string | Uint8Array | Uint8Array[]} data - File contents
   */
  addFile(name, data) {
    let bytes;
    if (typeof data === "string") {
      bytes = new TextEncoder().encode(data);
    } else if (Array.isArray(data)) {
      let totalLen = 0;
      for (const c of data) totalLen += c.length;
      bytes = new Uint8Array(totalLen);
      let offset = 0;
      for (const c of data) {
        bytes.set(c, offset);
        offset += c.length;
      }
    } else {
      bytes = data;
    }
    this.files.push({ name, data: bytes });
  }
  /**
   * Generate the ZIP file as a Blob.
   * @returns {Blob}
   */
  toBlob() {
    const parts = [];
    const centralDir = [];
    let localOffset = 0;
    for (const file of this.files) {
      const nameBytes = new TextEncoder().encode(file.name);
      const crc = this._crc32(file.data);
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const lhView = new DataView(localHeader.buffer);
      lhView.setUint32(0, 67324752, true);
      lhView.setUint16(4, 20, true);
      lhView.setUint16(6, 0, true);
      lhView.setUint16(8, 0, true);
      lhView.setUint16(10, 0, true);
      lhView.setUint16(12, 0, true);
      lhView.setUint32(14, crc, true);
      lhView.setUint32(18, file.data.length, true);
      lhView.setUint32(22, file.data.length, true);
      lhView.setUint16(26, nameBytes.length, true);
      lhView.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);
      parts.push(localHeader);
      parts.push(file.data);
      const cdEntry = new Uint8Array(46 + nameBytes.length);
      const cdView = new DataView(cdEntry.buffer);
      cdView.setUint32(0, 33639248, true);
      cdView.setUint16(4, 20, true);
      cdView.setUint16(6, 20, true);
      cdView.setUint16(8, 0, true);
      cdView.setUint16(10, 0, true);
      cdView.setUint16(12, 0, true);
      cdView.setUint16(14, 0, true);
      cdView.setUint32(16, crc, true);
      cdView.setUint32(20, file.data.length, true);
      cdView.setUint32(24, file.data.length, true);
      cdView.setUint16(28, nameBytes.length, true);
      cdView.setUint16(30, 0, true);
      cdView.setUint16(32, 0, true);
      cdView.setUint16(34, 0, true);
      cdView.setUint16(36, 0, true);
      cdView.setUint32(38, 0, true);
      cdView.setUint32(42, localOffset, true);
      cdEntry.set(nameBytes, 46);
      centralDir.push(cdEntry);
      localOffset += localHeader.length + file.data.length;
    }
    const cdStartOffset = localOffset;
    let cdSize = 0;
    for (const entry of centralDir) {
      parts.push(entry);
      cdSize += entry.length;
    }
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 101010256, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, this.files.length, true);
    eocdView.setUint16(10, this.files.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, cdStartOffset, true);
    eocdView.setUint16(20, 0, true);
    parts.push(eocd);
    return new Blob(parts, { type: "application/octet-stream" });
  }
  /**
   * CRC-32 calculation.
   * @param {Uint8Array} data
   * @returns {number}
   */
  _crc32(data) {
    if (!_BrowserZip._crc32Table) {
      _BrowserZip._crc32Table = new Uint32Array(256);
      for (let i = 0; i < 256; i++) {
        let c = i;
        for (let j = 0; j < 8; j++) {
          c = c & 1 ? 3988292384 ^ c >>> 1 : c >>> 1;
        }
        _BrowserZip._crc32Table[i] = c;
      }
    }
    const table = _BrowserZip._crc32Table;
    let crc = 4294967295;
    for (let i = 0; i < data.length; i++) {
      crc = table[(crc ^ data[i]) & 255] ^ crc >>> 8;
    }
    return (crc ^ 4294967295) >>> 0;
  }
};

// src/Formats.ts
var F = {
  THOUSANDS_SEP: "#,##0",
  CURRENCY_PLN: '#,##0.00 "z\u0142"',
  CURRENCY_EUR: "#,##0.00 \u20AC",
  PERCENTAGE: "0%",
  SCIENTIFIC: "0.00E+00",
  TWO_DECIMALS: "#,##0.00",
  TEXT: "@",
  LEADING_ZEROS: "000000000",
  DATE_SHORT: "dd.mm.yyyy",
  DATE_LONG: "d mmmm yyyy",
  DATE_DAY_MONTH_YEAR: "dd-mm-yyyy",
  DATE_ISO: "yyyy-mm-dd",
  DATE_MONTH_YEAR: "mmmm yyyy",
  DATE_WEEKDAY: "dddd, d mmmm yyyy",
  DATE_DAY_MONTH: "d mmmm",
  DATE_YEAR_ONLY: "yyyy",
  DATETIME_SHORT: "dd.mm.yyyy hh:mm",
  DATETIME_LONG: "d mmmm yyyy hh:mm:ss",
  TIME_HH_MM: "hh:mm",
  TIME_HH_MM_SS: "hh:mm:ss",
  TIME_12H: "h:mm AM/PM",
  DATETIME_24H: "dd.mm.yyyy hh:mm:ss",
  DATETIME_ISO: 'yyyy-mm-dd"T"hh:mm:ss',
  TIME_MS: "hh:mm:ss.000"
};
function isFormattedCell(val) {
  return val !== null && val !== void 0 && typeof val === "object" && "format" in val && typeof val.format === "string";
}
function unwrapCell(val) {
  if (isFormattedCell(val)) {
    return val.value;
  }
  return val;
}
function getFormat(val) {
  if (isFormattedCell(val)) {
    return val.format;
  }
  return null;
}

// src/xmlUtils.ts
function columnIndexToLetter(colIndex) {
  let n = colIndex + 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

// src/writerHelpers.ts
var OA_EPOCH_1900 = Date.UTC(1899, 11, 30);
var MARCH_1_1900 = Date.UTC(1900, 2, 1);
function timeToOaDate(timeMs, epochMs) {
  let serial = (timeMs - epochMs) / 864e5;
  if (epochMs === OA_EPOCH_1900 && timeMs < MARCH_1_1900) {
    serial -= 1;
  }
  return serial;
}
var MAX_ROWS = 1048576;
var MAX_COLS = 16384;
var MAX_CELL_CHARS = 32767;
function assertCellTextLength(text) {
  if (text.length > MAX_CELL_CHARS) {
    throw new Error(
      `Excel cell text exceeds ${MAX_CELL_CHARS} UTF-16 code units (got ${text.length}).`
    );
  }
}
function assertColumnCount(columnCount) {
  if (!Number.isInteger(columnCount) || columnCount < 1 || columnCount > MAX_COLS) {
    throw new Error(
      `Excel column count must be 1..${MAX_COLS} (got ${columnCount}).`
    );
  }
}
function assertRowFitsColumns(kind, length, columnCount) {
  if (length > columnCount) {
    throw new Error(
      `Excel ${kind} length ${length} exceeds columnCount ${columnCount}.`
    );
  }
}
function assertRowCount(totalRows) {
  if (totalRows > MAX_ROWS) {
    throw new Error(
      `Excel row count exceeds ${MAX_ROWS} (got ${totalRows}).`
    );
  }
}

// src/browser/BrowserXlsxWriter.ts
var BrowserXlsxWriter = class {
  constructor() {
    this._zip = new BrowserZip();
    this._sc = 0;
    this._sl = [];
    this._sstA = [];
    this._sstM = /* @__PURE__ */ new Map();
    this._sstCnt = 0;
    this._cw = [];
    this._afOn = false;
    this._oaE = Date.UTC(1899, 11, 30);
    this._csb = null;
    this._csrn = 0;
    this._csSC = 0;
    this._csEC = 0;
    this._csAF = false;
    this._isS = false;
    this._csCL = [];
    this._fmtReg = /* @__PURE__ */ new Map();
    this._fmtXf = /* @__PURE__ */ new Map();
    this._nextNfmt = 165;
    this._nextXf = 6;
  }
  _cl(i) {
    return columnIndexToLetter(i);
  }
  _san(n) {
    if (!n || typeof n !== "string") return `Sheet${this._sc + 1}`;
    let s = n.replace(/[\\/*?[\]:]/g, "_");
    if (s.length > 31) s = s.substring(0, 31);
    if (!s.trim()) s = `Sheet${this._sc + 1}`;
    return s;
  }
  _ex(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  }
  _esc(s) {
    let need = false;
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (c === 38 || c === 60 || c === 62 || c === 34 || c === 39 || c >= 0 && c <= 8 || c === 11 || c === 12 || c >= 14 && c <= 31) {
        need = true;
        break;
      }
    }
    if (!need) return s;
    const escaped = s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
    let out = "";
    for (let i = 0; i < escaped.length; i++) {
      const c = escaped.charCodeAt(i);
      if (c <= 8 || c === 11 || c === 12 || c >= 14 && c <= 31) continue;
      out += escaped[i];
    }
    return out;
  }
  _fmn(n) {
    return /[\s\-+=()!@#$%^&]/.test(n) || /^[0-9]/.test(n) ? `'${n.replace(/'/g, "''")}'` : n;
  }
  addSheet(name, hidden = false) {
    const sn = this._san(name);
    this._sc++;
    this._sl.push({ name: sn, path: `xl/worksheets/sheet${this._sc}.xml`, hidden, fn: `sheet${this._sc}.xml`, id: this._sc, rId: `rId${this._sc}`, fhr: null });
  }
  _wsc(bb, val, cr, rn, so) {
    assertCellTextLength(val);
    let idx = this._sstM.get(val);
    if (idx === void 0) {
      idx = this._sstA.length;
      this._sstA.push(val);
      this._sstM.set(val, idx);
    }
    this._sstCnt++;
    const sa = so !== void 0 ? ` s="${so}"` : "";
    bb.writeString(`<c r="${cr}${rn}" t="s"${sa}><v>${idx}</v></c>`);
  }
  _wscStyle(bb, val, cr, rn, styleId) {
    assertCellTextLength(val);
    let idx = this._sstM.get(val);
    if (idx === void 0) {
      idx = this._sstA.length;
      this._sstA.push(val);
      this._sstM.set(val, idx);
    }
    this._sstCnt++;
    bb.writeString(`<c r="${cr}${rn}" t="s" s="${styleId}"><v>${idx}</v></c>`);
  }
  _oa(d) {
    return timeToOaDate(d.getTime(), this._oaE);
  }
  _regFmt(fs) {
    const existing = this._fmtXf.get(fs);
    if (existing !== void 0) return existing;
    const nid = this._nextNfmt++;
    this._fmtReg.set(fs, nid);
    const xf = this._nextXf++;
    this._fmtXf.set(fs, xf);
    return xf;
  }
  _wcv(bb, raw, cr, rn) {
    if (raw === null || raw === void 0) return;
    const fs = getFormat(raw);
    const v = unwrapCell(raw);
    const xf = fs !== null ? this._regFmt(fs) : -1;
    const sa = xf >= 0 ? ` s="${xf}"` : "";
    if (v === null || v === void 0) return;
    if (typeof v === "number") {
      if (Number.isFinite(v)) bb.writeString(`<c r="${cr}${rn}"${sa}><v>${v}</v></c>`);
      else this._wsc(bb, v.toString(), cr, rn, xf >= 0 ? xf : void 0);
    } else if (typeof v === "bigint") this._wsc(bb, v.toString(), cr, rn, xf >= 0 ? xf : void 0);
    else if (typeof v === "boolean") bb.writeString(`<c r="${cr}${rn}" t="b"${sa}><v>${v ? 1 : 0}</v></c>`);
    else if (v instanceof Date) {
      const o = this._oa(v);
      if (Number.isFinite(o)) bb.writeString(`<c r="${cr}${rn}" s="${xf >= 0 ? xf : 1}"><v>${o}</v></c>`);
      else this._wsc(bb, v.toString(), cr, rn, xf >= 0 ? xf : void 0);
    } else this._wsc(bb, v.toString(), cr, rn, xf >= 0 ? xf : void 0);
  }
  _sheetHead(bb, cc, isFirst, hasAF) {
    bb.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    bb.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    bb.writeString('<dimension ref="A1"/>');
    if (hasAF) bb.writeString(`<sheetViews><sheetView ${isFirst ? 'tabSelected="1" ' : ""}workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>`);
    else bb.writeString(`<sheetViews><sheetView ${isFirst ? 'tabSelected="1" ' : ""}workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>`);
    bb.writeString("<cols>");
    for (let i = 0; i < cc; i++) {
      const w = this._cw[i] > 0 ? this._cw[i] : 10;
      bb.writeString(`<col min="${i + 1}" max="${i + 1}" width="${w}" bestFit="1" customWidth="1"/>`);
    }
    bb.writeString("</cols><sheetData>");
  }
  startSheet(name, colCount, headers, options = {}) {
    if (this._isS) throw new Error("Already streaming");
    assertColumnCount(colCount);
    if (headers !== void 0) {
      assertRowFitsColumns("header", headers.length, colCount);
      for (const h of headers) assertCellTextLength(h);
    }
    const { hidden = false, doAutofilter = true, headerStyle = "bold" } = options;
    let hs = 3;
    if (headerStyle === "fill") hs = 4;
    else if (headerStyle === "bold+fill") hs = 5;
    this.addSheet(name, hidden);
    this._isS = true;
    const csb = new BrowserBigBuffer();
    this._csb = csb;
    this._csrn = 0;
    this._csSC = 0;
    this._csEC = colCount;
    this._csAF = doAutofilter && headers !== void 0;
    this._cw = new Array(colCount).fill(-1);
    const cl = new Array(colCount);
    for (let i = 0; i < colCount; i++) cl[i] = this._cl(i);
    this._csCL = cl;
    if (headers) for (let i = 0; i < colCount; i++) {
      let w = 1.3 * (headers[i] ? headers[i].length : 0) + 3;
      if (w > 80) w = 80;
      if (this._cw[i] < w) this._cw[i] = w;
    }
    this._sheetHead(csb, colCount, this._sc === 1, this._csAF);
    if (headers) {
      this._csrn++;
      csb.writeString(`<row r="${this._csrn}">`);
      for (let c = 0; c < headers.length; c++) this._wscStyle(csb, headers[c], cl[c], this._csrn, hs);
      csb.writeString("</row>");
    }
  }
  writeRow(row) {
    if (!this._isS || this._csb === null) throw new Error("Not streaming");
    if (row.length !== this._csEC) throw new Error(`Row length mismatch. Expected ${this._csEC} columns, got ${row.length}`);
    const csb = this._csb;
    this._csrn++;
    assertRowCount(this._csrn);
    csb.writeString(`<row r="${this._csrn}">`);
    for (let c = 0; c < row.length; c++) this._wcv(csb, row[c], this._csCL[c], this._csrn);
    csb.writeString("</row>");
  }
  endSheet() {
    if (!this._isS || this._csb === null) throw new Error("Not streaming");
    const csb = this._csb;
    csb.writeString("</sheetData>");
    if (this._csAF && this._csEC > 0) {
      this._afOn = true;
      const fr = `A1:${this._csCL[this._csEC - 1]}${this._csrn}`;
      csb.writeString(`<autoFilter ref="${fr}"/>`);
      const sh = this._sl[this._sc - 1];
      sh.fhr = `${this._fmn(sh.name)}!$A$1:$${this._csCL[this._csEC - 1]}$${this._csrn}`;
    }
    csb.writeString("</worksheet>");
    this._zip.addFile(this._sl[this._sc - 1].path, csb.toUint8Array());
    this._isS = false;
    this._csb = null;
  }
  writeSheet(rows, headers = null, options = {}) {
    const doAutofilter = options.doAutofilter !== false;
    const headerStyle = options.headerStyle || "bold";
    let hs = 3;
    if (headerStyle === "fill") hs = 4;
    else if (headerStyle === "bold+fill") hs = 5;
    const bb = new BrowserBigBuffer();
    const cc = rows.length > 0 ? rows[0].length : headers ? headers.length : 0;
    if (cc > 0) assertColumnCount(cc);
    const tr = rows.length + (headers ? 1 : 0);
    assertRowCount(tr);
    if (headers) {
      assertRowFitsColumns("header", headers.length, cc);
      for (const h of headers) assertCellTextLength(h);
    }
    for (let r = 0; r < rows.length; r++) assertRowFitsColumns("row", rows[r].length, cc);
    this._cw = new Array(cc).fill(-1);
    const cl = new Array(cc);
    for (let i = 0; i < cc; i++) cl[i] = this._cl(i);
    if (headers) for (let i = 0; i < cc; i++) {
      let w = 1.3 * (headers[i] ? headers[i].length : 0) + 3;
      if (w > 80) w = 80;
      if (this._cw[i] < w) this._cw[i] = w;
    }
    for (let r = 0; r < Math.min(rows.length, 100); r++) for (let c = 0; c < rows[r].length; c++) {
      const raw = rows[r][c];
      if (raw == null) continue;
      const v = isFormattedCell(raw) ? raw.value : raw;
      if (v === null || v === void 0) continue;
      let w = 1.3 * (v instanceof Date ? 10 : v.toString().length) + 3;
      if (w > 80) w = 80;
      if (this._cw[c] < w) this._cw[c] = w;
    }
    this._sheetHead(bb, cc, this._sc === 1, doAutofilter && !!headers);
    let rn = 0;
    if (headers) {
      rn++;
      bb.writeString(`<row r="${rn}">`);
      for (let c = 0; c < headers.length; c++) this._wscStyle(bb, headers[c], cl[c], rn, hs);
      bb.writeString("</row>");
    }
    for (let r = 0; r < rows.length; r++) {
      rn++;
      bb.writeString(`<row r="${rn}">`);
      for (let c = 0; c < rows[r].length; c++) this._wcv(bb, rows[r][c], cl[c], rn);
      bb.writeString("</row>");
    }
    bb.writeString("</sheetData>");
    if (doAutofilter && headers && cc > 0) {
      this._afOn = true;
      const fr = `A1:${cl[cc - 1]}${tr}`;
      bb.writeString(`<autoFilter ref="${fr}"/>`);
      const sh = this._sl[this._sc - 1];
      sh.fhr = `${this._fmn(sh.name)}!$A$1:$${cl[cc - 1]}$${tr}`;
    }
    bb.writeString("</worksheet>");
    this._zip.addFile(this._sl[this._sc - 1].path, bb.toUint8Array());
  }
  finalize() {
    this._writeSst();
    this._writeStyles();
    this._writeWb();
    this._writeCt();
    this._writeRels();
    return this._zip.toBlob();
  }
  _writeSst() {
    const bb = new BrowserBigBuffer();
    bb.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    bb.writeString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this._sstCnt}" uniqueCount="${this._sstA.length}">`);
    for (const t of this._sstA) {
      const c = this._esc(t);
      if (c.length > 0 && (c[0] === " " || c[c.length - 1] === " " || /[\t\n\r]/.test(c))) bb.writeString(`<si><t xml:space="preserve">${c}</t></si>`);
      else bb.writeString(`<si><t>${c}</t></si>`);
    }
    bb.writeString("</sst>");
    this._zip.addFile("xl/sharedStrings.xml", bb.toUint8Array());
  }
  _writeStyles() {
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
      const ef = fs.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
      nfs += `<numFmt numFmtId="${nid}" formatCode="${ef}"/>`;
      nfc++;
    }
    const sf = [...this._fmtXf.entries()].sort((a, b) => a[1] - b[1]);
    for (const [fs] of sf) {
      const nid = this._fmtReg.get(fs);
      xfe.push(`<xf numFmtId="${nid}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`);
    }
    this._zip.addFile("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
${xfe.join("\n")}
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`);
  }
  _writeWb() {
    let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews><sheets>`;
    for (const s of this._sl) x += `<sheet name="${this._ex(s.name)}" sheetId="${s.id}"${s.hidden ? ' state="hidden"' : ""} r:id="${s.rId}"/>`;
    x += "</sheets>";
    if (this._afOn) {
      x += "<definedNames>";
      for (const s of this._sl) if (s.fhr) x += `<definedName name="_xlnm._FilterDatabase" localSheetId="${s.id - 1}" hidden="1">${this._esc(s.fhr)}</definedName>`;
      x += "</definedNames>";
    }
    x += '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>';
    this._zip.addFile("xl/workbook.xml", x);
  }
  _writeCt() {
    let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
    for (const s of this._sl) x += `<Override PartName="/${s.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    x += "</Types>";
    this._zip.addFile("[Content_Types].xml", x);
  }
  _writeRels() {
    this._zip.addFile("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    let r = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const s of this._sl) r += `<Relationship Id="${s.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.fn}"/>`;
    let n = this._sl.length + 1;
    r += `<Relationship Id="rId${n++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    r += `<Relationship Id="rId${n++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
    r += "</Relationships>";
    this._zip.addFile("xl/_rels/workbook.xml.rels", r);
  }
};

// src/browser/BrowserXlsbWriter.ts
var INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;
var BrowserXlsbWriter = class {
  constructor() {
    this._zip = new BrowserZip();
    this._sheetCount = 0;
    this._sheetList = [];
    this._sstDic = /* @__PURE__ */ new Map();
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
    this._fmtMap = /* @__PURE__ */ new Map();
    this._nextNid = 167;
    this._nextXf = 4;
    this._sheet1Bytes = new Uint8Array([
      129,
      1,
      0,
      147,
      1,
      23,
      203,
      4,
      2,
      0,
      64,
      0,
      0,
      0,
      0,
      0,
      0,
      255,
      255,
      255,
      255,
      255,
      255,
      255,
      255,
      0,
      0,
      0,
      0,
      148,
      1,
      16,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      133,
      1,
      0,
      137,
      1,
      30,
      220,
      3,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      64,
      0,
      0,
      0,
      100,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      152,
      1,
      36,
      3,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      138,
      1,
      0,
      134,
      1,
      0,
      37,
      6,
      1,
      0,
      2,
      14,
      0,
      128,
      149,
      8,
      2,
      5,
      0,
      38,
      0,
      229,
      3,
      12,
      255,
      255,
      255,
      255,
      8,
      0,
      44,
      1,
      0,
      0,
      0,
      0,
      145,
      1,
      0,
      37,
      6,
      1,
      0,
      2,
      14,
      0,
      128,
      128,
      8,
      2,
      5,
      0,
      38,
      0,
      0,
      25,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      44,
      1,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      7,
      12,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      146,
      1,
      0,
      151,
      4,
      66,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      221,
      3,
      2,
      16,
      0,
      220,
      3,
      48,
      102,
      102,
      102,
      102,
      102,
      102,
      230,
      63,
      102,
      102,
      102,
      102,
      102,
      102,
      230,
      63,
      0,
      0,
      0,
      0,
      0,
      0,
      232,
      63,
      0,
      0,
      0,
      0,
      0,
      0,
      232,
      63,
      51,
      51,
      51,
      51,
      51,
      51,
      211,
      63,
      51,
      51,
      51,
      51,
      51,
      51,
      211,
      63,
      37,
      6,
      1,
      0,
      0,
      16,
      0,
      128,
      128,
      24,
      16,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      38,
      0,
      130,
      1,
      0
    ]);
    this._workbookBinStart = new Uint8Array([
      131,
      1,
      0,
      128,
      1,
      50,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      2,
      0,
      0,
      0,
      120,
      0,
      108,
      0,
      1,
      0,
      0,
      0,
      55,
      0,
      1,
      0,
      0,
      0,
      54,
      0,
      5,
      0,
      0,
      0,
      50,
      0,
      52,
      0,
      51,
      0,
      50,
      0,
      54,
      0,
      153,
      1,
      12,
      32,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      37,
      6,
      1,
      0,
      3,
      15,
      0,
      128,
      151,
      16,
      52,
      24,
      0,
      0,
      0,
      67,
      0,
      58,
      0,
      92,
      0,
      115,
      0,
      113,
      0,
      108,
      0,
      115,
      0,
      92,
      0,
      84,
      0,
      101,
      0,
      115,
      0,
      116,
      0,
      121,
      0,
      90,
      0,
      97,
      0,
      112,
      0,
      105,
      0,
      115,
      0,
      117,
      0,
      88,
      0,
      108,
      0,
      115,
      0,
      98,
      0,
      92,
      0,
      38,
      0,
      37,
      6,
      1,
      0,
      0,
      16,
      0,
      128,
      129,
      24,
      130,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      47,
      0,
      0,
      0,
      49,
      0,
      51,
      0,
      95,
      0,
      110,
      0,
      99,
      0,
      114,
      0,
      58,
      0,
      49,
      0,
      95,
      0,
      123,
      0,
      49,
      0,
      54,
      0,
      53,
      0,
      48,
      0,
      56,
      0,
      68,
      0,
      54,
      0,
      57,
      0,
      45,
      0,
      67,
      0,
      70,
      0,
      56,
      0,
      55,
      0,
      45,
      0,
      52,
      0,
      55,
      0,
      54,
      0,
      57,
      0,
      45,
      0,
      56,
      0,
      52,
      0,
      53,
      0,
      54,
      0,
      45,
      0,
      68,
      0,
      52,
      0,
      65,
      0,
      52,
      0,
      48,
      0,
      49,
      0,
      49,
      0,
      51,
      0,
      49,
      0,
      53,
      0,
      54,
      0,
      55,
      0,
      125,
      0,
      47,
      0,
      0,
      0,
      47,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      38,
      0,
      135,
      1,
      0,
      37,
      6,
      1,
      0,
      2,
      16,
      0,
      128,
      128,
      24,
      16,
      0,
      0,
      0,
      0,
      13,
      0,
      0,
      0,
      255,
      255,
      255,
      255,
      0,
      0,
      0,
      0,
      38,
      0,
      158,
      1,
      29,
      0,
      0,
      0,
      0,
      158,
      22,
      0,
      0,
      180,
      105,
      0,
      0,
      232,
      38,
      0,
      0,
      88,
      2,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      120,
      136,
      1,
      0,
      143,
      1,
      0
    ]);
    this._workbookBinMiddle = new Uint8Array([144, 1, 0]);
    this._workbookBinEnd = new Uint8Array([
      157,
      1,
      26,
      53,
      234,
      2,
      0,
      1,
      0,
      0,
      0,
      100,
      0,
      0,
      0,
      252,
      169,
      241,
      210,
      77,
      98,
      80,
      63,
      1,
      0,
      0,
      0,
      106,
      0,
      155,
      1,
      1,
      0,
      35,
      4,
      3,
      15,
      0,
      0,
      171,
      16,
      1,
      1,
      36,
      0,
      132,
      1,
      0
    ]);
    this._stylesBin = new Uint8Array([
      150,
      2,
      0,
      231,
      4,
      4,
      2,
      0,
      0,
      0,
      44,
      44,
      164,
      0,
      19,
      0,
      0,
      0,
      121,
      0,
      121,
      0,
      121,
      0,
      121,
      0,
      92,
      0,
      45,
      0,
      109,
      0,
      109,
      0,
      92,
      0,
      45,
      0,
      100,
      0,
      100,
      0,
      92,
      0,
      32,
      0,
      104,
      0,
      104,
      0,
      58,
      0,
      109,
      0,
      109,
      0,
      44,
      30,
      166,
      0,
      12,
      0,
      0,
      0,
      121,
      0,
      121,
      0,
      121,
      0,
      121,
      0,
      92,
      0,
      45,
      0,
      109,
      0,
      109,
      0,
      92,
      0,
      45,
      0,
      100,
      0,
      100,
      0,
      232,
      4,
      0,
      227,
      4,
      4,
      2,
      0,
      0,
      0,
      43,
      39,
      220,
      0,
      0,
      0,
      144,
      1,
      0,
      0,
      0,
      2,
      0,
      0,
      7,
      1,
      0,
      0,
      0,
      0,
      0,
      255,
      2,
      7,
      0,
      0,
      0,
      67,
      0,
      97,
      0,
      108,
      0,
      105,
      0,
      98,
      0,
      114,
      0,
      105,
      0,
      43,
      39,
      220,
      0,
      1,
      0,
      188,
      2,
      0,
      0,
      0,
      2,
      238,
      0,
      7,
      1,
      0,
      0,
      0,
      0,
      0,
      255,
      2,
      7,
      0,
      0,
      0,
      67,
      0,
      97,
      0,
      108,
      0,
      105,
      0,
      98,
      0,
      114,
      0,
      105,
      0,
      37,
      6,
      1,
      0,
      2,
      14,
      0,
      128,
      129,
      8,
      0,
      38,
      0,
      228,
      4,
      0,
      219,
      4,
      4,
      2,
      0,
      0,
      0,
      45,
      68,
      0,
      0,
      0,
      0,
      3,
      64,
      0,
      0,
      0,
      0,
      0,
      255,
      3,
      65,
      0,
      0,
      255,
      255,
      255,
      255,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      45,
      68,
      17,
      0,
      0,
      0,
      3,
      64,
      0,
      0,
      0,
      0,
      0,
      255,
      3,
      65,
      0,
      0,
      255,
      255,
      255,
      255,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      220,
      4,
      0,
      229,
      4,
      4,
      1,
      0,
      0,
      0,
      46,
      51,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      230,
      4,
      0,
      242,
      4,
      4,
      1,
      0,
      0,
      0,
      47,
      16,
      255,
      255,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      0,
      0,
      243,
      4,
      0,
      233,
      4,
      4,
      6,
      // 6 xfs
      0,
      0,
      0,
      47,
      16,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      0,
      0,
      47,
      16,
      0,
      0,
      164,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      1,
      0,
      47,
      16,
      0,
      0,
      166,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      1,
      0,
      47,
      16,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      1,
      0,
      // index=3: bold
      47,
      16,
      0,
      0,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      1,
      0,
      // index=4: fill
      47,
      16,
      0,
      0,
      0,
      0,
      1,
      0,
      1,
      0,
      0,
      0,
      0,
      0,
      16,
      16,
      1,
      0,
      // index=5: bold+fill
      234,
      4,
      0,
      235,
      4,
      4,
      1,
      0,
      0,
      0,
      37,
      6,
      1,
      0,
      2,
      17,
      0,
      128,
      128,
      24,
      16,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      38,
      0,
      48,
      28,
      0,
      0,
      0,
      0,
      1,
      0,
      0,
      0,
      8,
      0,
      0,
      0,
      78,
      0,
      111,
      0,
      114,
      0,
      109,
      0,
      97,
      0,
      108,
      0,
      110,
      0,
      121,
      0,
      236,
      4,
      0,
      249,
      3,
      4,
      0,
      0,
      0,
      0,
      250,
      3,
      0,
      252,
      3,
      80,
      0,
      0,
      0,
      0,
      17,
      0,
      0,
      0,
      84,
      0,
      97,
      0,
      98,
      0,
      108,
      0,
      101,
      0,
      83,
      0,
      116,
      0,
      121,
      0,
      108,
      0,
      101,
      0,
      77,
      0,
      101,
      0,
      100,
      0,
      105,
      0,
      117,
      0,
      109,
      0,
      50,
      0,
      17,
      0,
      0,
      0,
      80,
      0,
      105,
      0,
      118,
      0,
      111,
      0,
      116,
      0,
      83,
      0,
      116,
      0,
      121,
      0,
      108,
      0,
      101,
      0,
      76,
      0,
      105,
      0,
      103,
      0,
      104,
      0,
      116,
      0,
      49,
      0,
      54,
      0,
      253,
      3,
      0,
      35,
      4,
      2,
      14,
      0,
      0,
      235,
      8,
      0,
      246,
      8,
      42,
      0,
      0,
      0,
      0,
      17,
      0,
      0,
      0,
      83,
      0,
      108,
      0,
      105,
      0,
      99,
      0,
      101,
      0,
      114,
      0,
      83,
      0,
      116,
      0,
      121,
      0,
      108,
      0,
      101,
      0,
      76,
      0,
      105,
      0,
      103,
      0,
      104,
      0,
      116,
      0,
      49,
      0,
      247,
      8,
      0,
      236,
      8,
      0,
      36,
      0,
      35,
      4,
      3,
      15,
      0,
      0,
      176,
      16,
      0,
      178,
      16,
      50,
      0,
      0,
      0,
      0,
      21,
      0,
      0,
      0,
      84,
      0,
      105,
      0,
      109,
      0,
      101,
      0,
      83,
      0,
      108,
      0,
      105,
      0,
      99,
      0,
      101,
      0,
      114,
      0,
      83,
      0,
      116,
      0,
      121,
      0,
      108,
      0,
      101,
      0,
      76,
      0,
      105,
      0,
      103,
      0,
      104,
      0,
      116,
      0,
      49,
      0,
      179,
      16,
      0,
      177,
      16,
      0,
      36,
      0,
      151,
      2,
      0
    ]);
    this._binaryIndexBin = new Uint8Array([
      42,
      24,
      0,
      0,
      0,
      0,
      32,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      149,
      2,
      0
    ]);
    this._rRkIntegerLowerLimit = -1 << 29;
    this._rRkIntegerUpperLimit = (1 << 29) - 1;
    this._autoFilterStartBytes = new Uint8Array([161, 1, 16]);
    this._autoFilterEndBytes = new Uint8Array([162, 1, 0]);
    this._stickHeaderA1bytes = new Uint8Array([
      151,
      1,
      29,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      240,
      63,
      1,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      2,
      0,
      0,
      0,
      3
    ]);
    this._magicFilterExcel2016Fix0 = new Uint8Array([225, 2, 0, 229, 2, 0, 234, 2]);
    this._magicFilterExcel2016Fix1 = new Uint8Array([
      39,
      70,
      33,
      0,
      0,
      0,
      0,
      255,
      0,
      0,
      0,
      15,
      0,
      0,
      0,
      95,
      0,
      70,
      0,
      105,
      0,
      108,
      0,
      116,
      0,
      101,
      0,
      114,
      0,
      68,
      0,
      97,
      0,
      116,
      0,
      97,
      0,
      98,
      0,
      97,
      0,
      115,
      0,
      101,
      0,
      15,
      0,
      0,
      0,
      59,
      255,
      0
    ]);
    this._magicFilterExcel2016Fix2 = new Uint8Array([0, 0, 0, 0, 255, 255, 255, 255]);
  }
  _sanitizeSheetName(name) {
    if (!name || typeof name !== "string") return `Sheet${this._sheetCount + 1}`;
    let s = name.replace(INVALID_SHEET_NAME_CHARS, "_");
    if (s.length > 31) s = s.substring(0, 31);
    if (s.trim().length === 0) s = `Sheet${this._sheetCount + 1}`;
    return s;
  }
  _registerFormat(fmtString) {
    const existing = this._fmtMap.get(fmtString);
    if (existing !== void 0) return existing.xf;
    const nid = this._nextNid++;
    const xf = this._nextXf++;
    this._fmtMap.set(fmtString, { nid, xf });
    return xf;
  }
  _buildStylesBin() {
    if (this._fmtMap.size === 0) return this._stylesBin;
    const base = this._stylesBin;
    let fontEndOff = -1;
    for (let i = 3; i < base.length - 3; i++) {
      if (base[i] === 232 && base[i + 1] === 4 && base[i + 2] === 0) {
        fontEndOff = i;
        break;
      }
    }
    let xfBeginOff = -1;
    for (let i = fontEndOff + 3; i < base.length - 2; i++) {
      if (base[i] === 233 && base[i + 1] === 4) {
        xfBeginOff = i;
        break;
      }
    }
    let xfEndOff = -1;
    for (let i = xfBeginOff; i < base.length - 3; i++) {
      if (base[i] === 234 && base[i + 1] === 4 && base[i + 2] === 0) {
        xfEndOff = i;
        break;
      }
    }
    let fillBeginOff = -1;
    for (let i = fontEndOff + 3; i < base.length - 2; i++) {
      if (base[i] === 227 && base[i + 1] === 4) {
        fillBeginOff = i;
        break;
      }
    }
    const tf = 2 + this._fmtMap.size;
    const tx = 4 + this._fmtMap.size;
    const p = [];
    function vlq(v) {
      const r2 = [];
      while (v >= 128) {
        r2.push(v & 127 | 128);
        v >>>= 7;
      }
      r2.push(v & 127);
      return r2;
    }
    function brt(t, d) {
      return new Uint8Array([...vlq(t), ...vlq(d.length), ...d]);
    }
    function stFmt(ifmt, fs) {
      const cch = fs.length, rem = 2 + 4 + cch * 2;
      const buf = new Uint8Array(2 + rem);
      buf[0] = 44;
      buf[1] = rem;
      const dv = new DataView(buf.buffer);
      dv.setUint16(2, ifmt, true);
      dv.setUint32(4, cch, true);
      for (let i = 0; i < cch; i++) {
        const cp = fs.charCodeAt(i);
        buf[8 + i * 2] = cp & 255;
        buf[8 + i * 2 + 1] = cp >> 8 & 255;
      }
      return buf;
    }
    function xfRec(fid, ifmt, fl, be = 0) {
      const buf = new Uint8Array(18);
      const dv = new DataView(buf.buffer);
      buf[0] = 47;
      buf[1] = 16;
      dv.setUint16(2, fid, true);
      dv.setUint16(4, ifmt, true);
      buf[6] = be;
      buf[14] = 16;
      buf[15] = 16;
      dv.setUint16(16, fl, true);
      return buf;
    }
    p.push(base.subarray(0, 3));
    const fh = new Uint8Array(4);
    new DataView(fh.buffer).setUint16(0, tf, true);
    p.push(brt(615, fh));
    p.push(stFmt(164, "yyyy\\-mm\\-dd\\ hh:mm:ss"));
    p.push(stFmt(166, "yyyy\\-mm\\-dd"));
    for (const [fs, st] of this._fmtMap) p.push(stFmt(st.nid, fs));
    p.push(brt(616, new Uint8Array(0)));
    p.push(base.subarray(fillBeginOff, xfBeginOff));
    const xh = new Uint8Array(4);
    new DataView(xh.buffer).setUint16(0, tx, true);
    p.push(brt(617, xh));
    p.push(xfRec(0, 0, 0));
    p.push(xfRec(0, 164, 1));
    p.push(xfRec(0, 166, 1));
    p.push(xfRec(1, 0, 0, 1));
    const sortedXf = [...this._fmtMap.entries()].sort((a, b) => a[1].xf - b[1].xf);
    for (const [, st] of sortedXf) p.push(xfRec(0, st.nid, 1));
    p.push(brt(618, new Uint8Array(0)));
    p.push(base.subarray(xfEndOff + 3));
    let tl = 0;
    for (const c of p) tl += c.length;
    const r = new Uint8Array(tl);
    let pos = 0;
    for (const c of p) {
      r.set(c, pos);
      pos += c.length;
    }
    return r;
  }
  addSheet(sheetName, hidden = false) {
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
  startSheet(sheetName, columnCount, headers, options = {}) {
    if (this._isStreaming) throw new Error("Already streaming");
    assertColumnCount(columnCount);
    if (headers !== void 0) {
      assertRowFitsColumns("header", headers.length, columnCount);
      for (const h of headers) assertCellTextLength(h);
    }
    const { hidden = false, doAutofilter = true, headerStyle = "bold" } = options;
    let hs = 3;
    if (headerStyle === "fill") hs = 4;
    else if (headerStyle === "bold+fill") hs = 5;
    this.addSheet(sheetName, hidden);
    this._isStreaming = true;
    this._currentSheetBuffer = new BrowserBigBuffer();
    this._currentSheetRowNum = 0;
    this._currentSheetStartCol = 0;
    this._currentSheetEndCol = columnCount;
    this._currentSheetDoAutofilter = doAutofilter && headers !== void 0;
    this._colWidths = new Array(columnCount).fill(-1);
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
  writeRow(row) {
    const bb = this._currentSheetBuffer;
    if (!this._isStreaming || bb === null) throw new Error("Not streaming");
    if (row.length !== this._currentSheetEndCol - this._currentSheetStartCol) {
      throw new Error(`Row length mismatch. Expected ${this._currentSheetEndCol - this._currentSheetStartCol} columns, got ${row.length}`);
    }
    assertRowCount(this._currentSheetRowNum + 1);
    this.createRowHeader(bb, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
    for (let c = 0; c < row.length; c++) this._writeCell(bb, row[c], c);
    this._currentSheetRowNum++;
  }
  endSheet() {
    const bb = this._currentSheetBuffer;
    if (!this._isStreaming || bb === null) throw new Error("Not streaming");
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
        startRow: 0,
        endRow: this._currentSheetRowNum - 1,
        startColumn: this._currentSheetStartCol,
        endColumn: this._currentSheetEndCol - 1
      };
    }
    bb.write(this._sheet1Bytes.subarray(290));
    this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
    this._isStreaming = false;
    this._currentSheetBuffer = null;
  }
  writeSheet(rows, headers = null, options = {}) {
    const doAutofilter = options.doAutofilter !== false;
    const headerStyle = options.headerStyle || "bold";
    let hs = 3;
    if (headerStyle === "fill") hs = 4;
    else if (headerStyle === "bold+fill") hs = 5;
    const bb = new BrowserBigBuffer();
    const colCount = rows.length > 0 ? rows[0].length : headers ? headers.length : 0;
    if (colCount > 0) assertColumnCount(colCount);
    const totalBr = rows.length + (headers ? 1 : 0);
    assertRowCount(totalBr);
    if (headers) {
      assertRowFitsColumns("header", headers.length, colCount);
      for (const h of headers) assertCellTextLength(h);
    }
    for (let r = 0; r < rows.length; r++) assertRowFitsColumns("row", rows[r].length, colCount);
    this._colWidths = new Array(colCount).fill(-1);
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
        const raw = rows[r][c];
        if (raw == null) continue;
        const v = isFormattedCell(raw) ? raw.value : raw;
        if (v === null || v === void 0) continue;
        const len = v instanceof Date ? 10 : v.toString().length;
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
      rView.setInt32(0, 0, true);
      rView.setInt32(4, rn - 1, true);
      bb.write(rBuf);
      const cBuf = new Uint8Array(8), cView = new DataView(cBuf.buffer);
      cView.setInt32(0, 0, true);
      cView.setInt32(4, colCount - 1, true);
      bb.write(cBuf);
      bb.write(this._autoFilterEndBytes);
      this._sheetList[this._sheetCount - 1].filterData = {
        startRow: 0,
        endRow: rn,
        startColumn: 0,
        endColumn: colCount - 1
      };
    }
    bb.write(this._sheet1Bytes.subarray(290));
    this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
  }
  _writeSheetHeader(bb, startCol, endCol, doAf) {
    const sh = new Uint8Array(this._sheet1Bytes);
    const shView = new DataView(sh.buffer);
    shView.setInt32(40, startCol, true);
    shView.setInt32(44, endCol, true);
    if (this._sheetCount !== 1) sh[54] = 156;
    bb.write(sh.subarray(0, 84));
    if (doAf) bb.write(this._stickHeaderA1bytes);
    bb.write(sh.subarray(84, 159));
    bb.writeByte(134);
    bb.writeByte(3);
    for (let i = startCol; i < endCol; i++) {
      bb.writeByte(0);
      bb.writeByte(60);
      bb.writeByte(18);
      bb.writeInt32LE(i);
      bb.writeInt32LE(i);
      const w = this._colWidths[i] > 0 ? this._colWidths[i] : 10;
      bb.writeInt32LE(Math.round(w * 256));
      bb.writeByte(0);
      bb.writeByte(0);
      bb.writeByte(0);
      bb.writeByte(0);
      bb.writeByte(2);
    }
    bb.writeByte(0);
    bb.writeByte(135);
    bb.writeByte(3);
    bb.writeByte(0);
    bb.write(sh.subarray(159, 175));
    bb.write(new Uint8Array([38, 0]));
  }
  _writeCell(bb, raw, c) {
    if (raw === null || raw === void 0) return;
    const fmtString = getFormat(raw);
    const val = unwrapCell(raw);
    const styleNum = fmtString !== null ? this._registerFormat(fmtString) : 0;
    if (typeof val === "number") {
      if (Number.isInteger(val) && val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
        this.writeRkNumberInteger(bb, val, c, styleNum);
      } else {
        this.writeDouble(bb, val, c, styleNum);
      }
    } else if (typeof val === "bigint") {
      this.writeString(bb, val.toString(), c);
    } else if (typeof val === "boolean") {
      this.writeBool(bb, val, c);
    } else if (val instanceof Date) {
      const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
      this.writeDouble(bb, oaDate, c, fmtString !== null ? styleNum : 1);
    } else if (val !== null && val !== void 0) {
      this.writeString(bb, val.toString(), c, 0, fmtString !== null ? styleNum : void 0);
    }
  }
  createRowHeader(bb, rn, sc, ec) {
    bb.ensureCapacity(27);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(25);
    bb.writeUnsafeInt32LE(rn);
    bb.writeUnsafeInt32LE(0);
    bb.writeUnsafeByte(44);
    bb.writeUnsafeByte(1);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(1);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeInt32LE(sc);
    bb.writeUnsafeInt32LE(ec);
  }
  writeRkNumberInteger(bb, val, colNum, styleNum = 0) {
    bb.ensureCapacity(14);
    bb.writeUnsafeByte(2);
    bb.writeUnsafeByte(12);
    bb.writeUnsafeInt32LE(colNum);
    bb.writeUnsafeByte(styleNum);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeInt32LE(val << 2 | 2);
  }
  writeDouble(bb, val, colNum, styleNum = 0) {
    bb.ensureCapacity(18);
    bb.writeUnsafeByte(5);
    bb.writeUnsafeByte(16);
    bb.writeUnsafeInt32LE(colNum);
    bb.writeUnsafeByte(styleNum);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeDoubleLE(val);
  }
  writeBool(bb, val, colNum) {
    bb.ensureCapacity(13);
    bb.writeUnsafeByte(4);
    bb.writeUnsafeByte(9);
    bb.writeUnsafeInt32LE(colNum);
    bb.writeUnsafeInt32LE(0);
    bb.writeUnsafeByte(val ? 1 : 0);
  }
  writeDateTime(bb, date, colNum) {
    this.writeDouble(bb, timeToOaDate(date.getTime(), this._oaEpoch), colNum, 1);
  }
  writeString(bb, val, colNum, styleIndex = 0, styleOverride) {
    assertCellTextLength(val);
    let index = this._sstDic.get(val);
    if (index === void 0) {
      index = this._sstCntUnique++;
      this._sstDic.set(val, index);
    }
    this._sstCntAll++;
    const finalStyle = styleOverride !== void 0 ? styleOverride : styleIndex;
    bb.ensureCapacity(17);
    bb.writeUnsafeByte(7);
    bb.writeUnsafeByte(12);
    bb.writeUnsafeInt32LE(colNum);
    bb.writeUnsafeByte(finalStyle);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeByte(0);
    bb.writeUnsafeInt32LE(index);
  }
  _saveSst() {
    const bb = new BrowserBigBuffer();
    bb.writeByte(159);
    bb.writeByte(1);
    bb.writeByte(8);
    bb.writeInt32LE(this._sstCntAll);
    bb.writeInt32LE(this._sstCntUnique);
    for (const [txt] of this._sstDic) {
      const tl = txt.length;
      bb.writeByte(19);
      const rl = 5 + 2 * tl;
      if (rl >= 128) {
        bb.writeByte(128 + rl % 128);
        const tmp = rl >> 7;
        if (tmp >= 256) bb.writeByte(128 + tmp % 128);
        else bb.writeByte(tmp);
        bb.writeByte(rl >> 14);
        if (rl >> 14 > 0) bb.writeByte(0);
      } else {
        bb.writeByte(rl & 255);
        bb.writeByte(rl >> 8 & 255);
      }
      bb.writeInt32LE(tl);
      bb.writeUtf16LE(txt);
    }
    bb.writeByte(160);
    bb.writeByte(1);
    bb.writeByte(0);
    this._zip.addFile("xl/sharedStrings.bin", bb.toUint8Array());
  }
  _writeFilterDefinedName(wbBuffers, sheet, sheetNum) {
    const fd = sheet.filterData;
    const fix1 = new Uint8Array(this._magicFilterExcel2016Fix1);
    fix1[7] = sheet.sheetId - 1;
    fix1[fix1.length - 2] = sheetNum;
    wbBuffers.push(fix1);
    const rBuf = new Uint8Array(8), rView = new DataView(rBuf.buffer);
    rView.setInt32(0, fd.startRow, true);
    rView.setInt32(4, fd.endRow, true);
    wbBuffers.push(rBuf);
    const cBuf = new Uint8Array(4), cView = new DataView(cBuf.buffer);
    cView.setInt16(0, fd.startColumn, true);
    cView.setInt16(2, fd.endColumn, true);
    wbBuffers.push(cBuf);
    wbBuffers.push(this._magicFilterExcel2016Fix2);
  }
  finalize() {
    this._saveSst();
    this._zip.addFile("xl/styles.bin", this._buildStylesBin());
    const wbBuffers = [this._workbookBinStart];
    for (const sheet of this._sheetList) {
      const rId = `rId${sheet.sheetId}`;
      const sn = sheet.name;
      const rl = 4 + 12 + sn.length * 2 + rId.length * 2;
      const buf = new Uint8Array(3 + rl);
      const view = new DataView(buf.buffer);
      buf[0] = 156;
      buf[1] = 1;
      buf[2] = rl;
      let pos = 3;
      view.setInt32(pos, sheet.hidden ? 1 : 0, true);
      pos += 4;
      view.setInt32(pos, sheet.sheetId, true);
      pos += 4;
      view.setInt32(pos, rId.length, true);
      pos += 4;
      for (let i = 0; i < rId.length; i++) {
        const c = rId.charCodeAt(i);
        buf[pos++] = c & 255;
        buf[pos++] = c >> 8 & 255;
      }
      view.setInt32(pos, sn.length, true);
      pos += 4;
      for (let i = 0; i < sn.length; i++) {
        const c = sn.charCodeAt(i);
        buf[pos++] = c & 255;
        buf[pos++] = c >> 8 & 255;
      }
      wbBuffers.push(buf);
    }
    wbBuffers.push(this._workbookBinMiddle);
    if (this._autofilterIsOn) {
      const filtered = this._sheetList.filter((s) => s.filterData !== void 0);
      if (filtered.length > 0) {
        wbBuffers.push(this._magicFilterExcel2016Fix0);
        const cnt = filtered.length;
        const firstByte = cnt <= 20 ? 16 + (cnt - 1) * 12 : 128 + (cnt - 21) * 12;
        wbBuffers.push(cnt <= 10 ? new Uint8Array([firstByte, cnt, 0, 0, 0]) : new Uint8Array([firstByte, Math.floor((cnt - 1) / 10), cnt, 0, 0, 0]));
        for (let i = 0; i < filtered.length; i++) {
          const si = filtered[i].sheetId - 1;
          const ib = new Uint8Array(12);
          ib[4] = si;
          ib[8] = si;
          wbBuffers.push(ib);
        }
        wbBuffers.push(new Uint8Array([226, 2, 0]));
        for (let i = 0; i < filtered.length; i++) this._writeFilterDefinedName(wbBuffers, filtered[i], i);
      }
    }
    wbBuffers.push(this._workbookBinEnd);
    let totalWbLen = 0;
    for (const b of wbBuffers) totalWbLen += b.length;
    const wbFinal = new Uint8Array(totalWbLen);
    let wbOffset = 0;
    for (const b of wbBuffers) {
      wbFinal.set(b, wbOffset);
      wbOffset += b.length;
    }
    this._zip.addFile("xl/workbook.bin", wbFinal);
    for (const s of this._sheetList) this._zip.addFile(`xl/worksheets/binaryIndex${s.sheetId}.bin`, this._binaryIndexBin);
    let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`;
    for (const s of this._sheetList) {
      ct += `<Override PartName="/${s.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>
<Override PartName="/xl/worksheets/binaryIndex${s.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>
`;
    }
    ct += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>`;
    this._zip.addFile("[Content_Types].xml", ct);
    let wr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
`;
    for (const s of this._sheetList) wr += `<Relationship Id="rId${s.sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.nameInArchive}"/>
`;
    wr += `<Relationship Id="rId${this._sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this._sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
</Relationships>`;
    this._zip.addFile("xl/_rels/workbook.bin.rels", wr);
    this._zip.addFile("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>`);
    for (const s of this._sheetList) {
      this._zip.addFile(`xl/worksheets/_rels/${s.nameInArchive}.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${s.sheetId}.bin"/>
</Relationships>`);
    }
    return this._zip.toBlob();
  }
};

// src/browser/download.ts
function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5e3);
}
function downloadXlsx(fileName, rows, headers, sheetName = "Sheet1") {
  const w = new BrowserXlsxWriter();
  w.addSheet(sheetName);
  w.writeSheet(rows, headers, { doAutofilter: true });
  triggerDownload(w.finalize(), fileName);
}
function downloadXlsxMultiSheet(fileName, sheets) {
  const w = new BrowserXlsxWriter();
  for (const s of sheets) {
    w.addSheet(s.name);
    w.writeSheet(s.rows, s.headers, { doAutofilter: true });
  }
  triggerDownload(w.finalize(), fileName);
}
function downloadXlsb(fileName, rows, headers, sheetName = "Sheet1") {
  const w = new BrowserXlsbWriter();
  w.addSheet(sheetName);
  w.writeSheet(rows, headers, { doAutofilter: true });
  triggerDownload(w.finalize(), fileName);
}
function downloadXlsbMultiSheet(fileName, sheets) {
  const w = new BrowserXlsbWriter();
  for (const s of sheets) {
    w.addSheet(s.name);
    w.writeSheet(s.rows, s.headers, { doAutofilter: true });
  }
  triggerDownload(w.finalize(), fileName);
}
export {
  BrowserBigBuffer,
  BrowserXlsbWriter,
  BrowserXlsxWriter,
  BrowserZip,
  F,
  downloadXlsb,
  downloadXlsbMultiSheet,
  downloadXlsx,
  downloadXlsxMultiSheet
};
