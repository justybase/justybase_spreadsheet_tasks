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
      this.chunks.push(this.currentBuffer.subarray(0, this.cursor));
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
    if (this.cursor + str.length <= this.chunkSize) {
      let ascii = true;
      for (let i = 0; i < str.length; i++) {
        if (str.charCodeAt(i) >= 128) {
          ascii = false;
          break;
        }
      }
      if (ascii) {
        for (let i = 0; i < str.length; i++) {
          this.currentBuffer[this.cursor++] = str.charCodeAt(i);
        }
        return;
      }
    }
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
      this.chunks.push(this.currentBuffer.subarray(0, this.cursor));
      this.currentBuffer = new Uint8Array(this.chunkSize);
      this.dataView = new DataView(this.currentBuffer.buffer);
      this.cursor = 0;
    }
    return this.chunks;
  }
  /** Overwrite a previously written 32-bit integer at an absolute offset. */
  patchInt32LE(offset, value) {
    let position = 0;
    for (const chunk of this.chunks) {
      if (offset >= position && offset + 4 <= position + chunk.length) {
        new DataView(chunk.buffer, chunk.byteOffset, chunk.byteLength).setInt32(offset - position, value, true);
        return true;
      }
      position += chunk.length;
    }
    if (offset >= position && offset + 4 <= position + this.cursor) {
      this.dataView.setInt32(offset - position, value, true);
      return true;
    }
    return false;
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

// src/xmlUtils.ts
function hasInvalidControlChars(str) {
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c < 32 && c !== 9 && c !== 10 && c !== 13) {
      return true;
    }
  }
  return false;
}
function stripInvalidControlChars(str) {
  let out = "";
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c === 9 || c === 10 || c === 13 || c >= 32 && c <= 55295 || c >= 57344 && c <= 65533) {
      out += str[i];
    }
  }
  return out;
}
function escapeXmlText(str) {
  if (!/[&<>"'\r]/.test(str) && !hasInvalidControlChars(str)) {
    return str;
  }
  let out = str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;").replace(/\r/g, "&#13;");
  if (hasInvalidControlChars(out)) {
    out = stripInvalidControlChars(out);
  }
  return out;
}
function escapeXlsxText(str) {
  return escapeXmlText(str.replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_"));
}
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

// src/Formats.ts
var ExcelError = class {
  constructor(code) {
    this.code = code;
  }
  toString() {
    return this.code;
  }
};
var RichText = class {
  constructor(runs) {
    this.runs = runs;
  }
  toString() {
    return this.runs.map((run) => run.text).join("");
  }
};
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
function isRichText(val) {
  return val instanceof RichText;
}
function isStyledCell(val) {
  return val !== null && val !== void 0 && typeof val === "object" && "style" in val && "value" in val && !("formula" in val);
}
function isFormulaCell(val) {
  return val !== null && val !== void 0 && typeof val === "object" && "formula" in val;
}
function getCellStyle(val) {
  if (isFormattedCell(val)) return { numFmt: val.format };
  if (isStyledCell(val)) return val.style;
  if (isFormulaCell(val)) return val.style ?? null;
  return null;
}
function getFormula(val) {
  return isFormulaCell(val) ? val.formula : null;
}
function unwrapCell(val) {
  if (isFormattedCell(val) || isStyledCell(val)) {
    return val.value;
  }
  if (isFormulaCell(val)) {
    return val.value ?? null;
  }
  return val;
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
function dateHasTime(value) {
  return value.getUTCHours() !== 0 || value.getUTCMinutes() !== 0 || value.getUTCSeconds() !== 0 || value.getUTCMilliseconds() !== 0;
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
function initColWidths(columnCount) {
  return new Array(columnCount).fill(-1);
}
var DEFAULT_COLUMN_WIDTH = 8.43;
var MAX_AUTOFIT_WIDTH = 200;
var COLUMN_WIDTH_PADDING = 5 / 7;
var AUTOFILTER_BUTTON_WIDTH = 2.4;
var AUTOFIT_SAMPLE_ROWS = 200;
function resolveAutofitRows(value) {
  if (value === void 0) return AUTOFIT_SAMPLE_ROWS;
  if (!Number.isInteger(value) || value < 0) {
    throw new RangeError(`autofitRows must be a non-negative integer, got ${value}`);
  }
  return value;
}
var AUTOFIT_FACTOR = 1.2;
var LONG_TEXT_FACTOR = 29 / 28;
var BOLD_WIDTH_FACTOR = 1.08;
var ITALIC_WIDTH_FACTOR = 1.05;
var DEFAULT_FONT_SIZE = 11;
var INDENT_WIDTH = 9 / 7;
function makeWidthCosts(groups) {
  const costs = /* @__PURE__ */ Object.create(null);
  for (const [characters, cost] of groups) {
    for (const character of characters) costs[character] = cost;
  }
  return costs;
}
var NARROW_FONT_COSTS = makeWidthCosts([
  ["@W", 1.75],
  ["%M", 1.7],
  ["m", 1.7],
  ["OQwGHN", 1.45],
  ["CD", 1.4],
  ["&U", 1.35],
  ["R", 1.25],
  ["BKAPT", 1.2],
  ["ESXbdpqu", 1.15],
  ["$0123456789Yhno", 1.1],
  ["#+<=>^`ae~", 1.1],
  ["?FZcg", 1.05],
  ["*_kLsvy|", 1],
  ["xz", 0.95],
  ["J", 0.9],
  ['"/', 0.85],
  ["\\r", 0.8],
  ["!-t", 0.75],
  ["()[]f{}", 0.7],
  ["I", 0.65],
  [",.:;", 0.65],
  ["l", 0.6],
  [" ij", 0.6],
  ["'", 0.5]
]);
var WIDE_FONT_COSTS = makeWidthCosts([
  ["@W", 2.5],
  ["M", 2.35],
  ["%", 2.3],
  ["m", 2.25],
  ["w", 2.05],
  ["&", 1.95],
  ["OQ", 1.9],
  ["NU", 1.85],
  ["GHD", 1.75],
  ["#K", 1.7],
  ["A", 1.7],
  ["+<=>^~", 1.65],
  ["V", 1.65],
  ["C", 1.6],
  ["BRXY", 1.6],
  ["Pbdhnopqu", 1.5],
  ["$0123456789e", 1.45],
  ["ESTZ", 1.45],
  ["*_ag", 1.4],
  ["FJLckvx|?`y{}", 1.35],
  ['"/\\', 1.3],
  ["s", 1.25],
  [" !'(),-.:;I[]fijlrtz", 1.25]
]);
var NARROW_FONT_NAMES = /* @__PURE__ */ new Set([
  "calibri",
  "calibri light",
  "aptos",
  "aptos narrow",
  "arial narrow"
]);
function usesNarrowFont(name) {
  if (!name) return true;
  return NARROW_FONT_NAMES.has(name.trim().toLowerCase());
}
function isCjkWideCodePoint(cp) {
  return cp >= 4352 && cp <= 4447 || // Hangul Jamo
  cp >= 11904 && cp <= 12350 || // CJK radicals and punctuation
  cp >= 12353 && cp <= 13311 || // Hiragana, Katakana, CJK compatibility
  cp >= 13312 && cp <= 19903 || // CJK extension A
  cp >= 19968 && cp <= 40959 || // CJK unified ideographs
  cp >= 40960 && cp <= 42191 || // Yi
  cp >= 44032 && cp <= 55203 || // Hangul syllables
  cp >= 63744 && cp <= 64255 || // CJK compatibility ideographs
  cp >= 65040 && cp <= 65049 || // vertical forms
  cp >= 65072 && cp <= 65135 || // CJK compatibility forms
  cp >= 65280 && cp <= 65376 || // fullwidth forms
  cp >= 65504 && cp <= 65510 || // fullwidth signs
  cp >= 131072 && cp <= 262141;
}
function isEmojiCodePoint(cp) {
  return cp >= 126976 && cp <= 129791 || cp >= 8960 && cp <= 9215 || cp >= 9728 && cp <= 10175 || cp >= 11008 && cp <= 11263;
}
function isZeroWidthCodePoint(cp) {
  return cp < 32 || cp >= 127 && cp <= 159 || cp >= 768 && cp <= 879 || cp >= 6832 && cp <= 6911 || cp >= 7616 && cp <= 7679 || cp >= 8400 && cp <= 8447 || cp >= 65056 && cp <= 65071;
}
var WIDE_UNICODE_COSTS = makeWidthCosts([
  ["\xA5\xC8\xC9\xCA\xCB\xDE\xE8\xE9\xEA\xEB\xF0\xF1\xF2\xF3\xF4\xF5\xF6\xF8\xF9\xFA\xFB\xFC\xFE\u0112\u0113\u0114\u0115\u0116\u0117\u0118\u0119\u011A\u011B\u0125\u0133\u0144\u0146\u0148\u014B\u014D\u014F\u0151\u015A\u015C\u015E\u0160\u0162\u0164\u0166\u0169\u016B\u016D\u016F\u0171\u0173\u0179\u017B\u017D\u0183\u018C\u018D\u018E\u019E\u01A5\u01A7\u01B4\u01BB\u01BC\u01BF\u01C9\u01D2\u01D4\u01D6\u01D8\u01DA\u01DC\u01DD\u01EB\u01ED\u01F9\u01FF\u0204\u0205\u0206\u0207\u020D\u020F\u0215\u0217\u0218\u021A\u021F\u0222\u0223\u0228\u0229\u022B\u022D\u022F\u0231\u0232\u023B\u023E\u0246\u024B", 1.35],
  ["\u0127\u0185\u01AC\u01E8", 1.4],
  ["\xA3\xAB\xB5\xBB\xC7\xDD\xDF\u0106\u0108\u010A\u010C\u0111\u013F\u0154\u0156\u0158\u0176\u0178\u0180\u0182\u0184\u0186\u018B\u0198\u01A6\u01B3\u01E5\u0210\u0212\u0243\u024C", 1.45],
  ["\xAC\xB1\xD7\xF7\u0194\u01C2\u0245\u024E", 1.5],
  ["\xB6\xC0\xC1\xC2\xC3\xC4\xC5\u0100\u0102\u0104\u010F\u0136\u0187\u01CD\u01DE\u01E0\u01F7\u01FA\u0200\u0202\u0221\u0226\u023A", 1.55],
  ["\xD0\u010E\u0110\u011C\u011E\u0120\u0122\u0124\u0149\u014A\u0189\u01A1\u01A4\u01E6\u01F4\u021E\u0220\u0235\u024A", 1.6],
  ["\xD1\xD2\xD3\xD4\xD5\xD6\xD8\xD9\xDA\xDB\xDC\u0126\u0143\u0145\u0147\u014C\u014E\u0150\u0168\u016A\u016C\u016E\u0170\u0172\u0181\u018F\u0193\u019D\u019F\u01B0\u01B1\u01B2\u01C8\u01D1\u01D3\u01D5\u01D7\u01D9\u01DB\u01E4\u01EA\u01EC\u01F8\u01FE\u020C\u020E\u0214\u0216\u022A\u022C\u022E\u0230\u0244", 1.7],
  ["\xAE\u01A3", 1.8],
  ["\u018A", 1.85],
  ["\u0175\u01A0\u01AF", 1.9],
  ["\u01E2", 1.95],
  ["\xBC\xBD\xBE\u0195\u01A2\u01C7\u01CC\u01E3", 2],
  ["\u0132\u03E3\u0416\u04C1\u04DC\u0502\u0503\u050B\u0906\u0911\u0912\u0913\u0914\u0974\u0975\u0E33\u1E3F\u1E41\u1E43\u20A5\u2106\u210B\u21FF\u220F\u226A\u226B", 2.05],
  ["\xC6\xE6\u01FC\u01FD\u0238\u0239\u03D6\u046D\u0518\u2105\u2154\u25EF", 2.1],
  ["\xA9\u205B\u2117", 2.15],
  ["\u0152\u0153\u039C\u041C\u1E3E\u1E40\u1E42\u2180\u2182\u21AD\u21AE\u21CE\u21D4\u21DC\u21DD\u221E", 2.2],
  ["\u019C\u03E2\u0409\u040A\u0428\u042E\u0468\u0496\u04CD\u0514\u20A8\u2120\u2188", 2.25],
  ["\u0174\u01CB\u01F6\u0429\u04A6\u0508\u050A\u051C\u0520\u0522\u1E80\u1E82\u1E84\u1E86\u1E88\u2014\u2015\u2162\u2177\u2190\u2191\u2192\u2193\u2196\u2197\u2198\u2199\u222D", 2.3],
  ["\u01C6\u01F3\u046C\u0479\u047C\u20A9\u2133\u214F", 2.35],
  ["\u2230", 2.4],
  ["\u01CA\u2152", 2.45],
  ["\u0460\u047E", 2.5],
  ["\u01C5\u01F2\u052B\u2001\u2003\u25FD\u25FE", 2.55],
  ["\u0478\u2116\u2121\u213B\u22D8\u22D9", 2.65],
  ["\u2166\u216B", 2.7],
  ["\u01C4\u01F1\u20AF", 2.75],
  ["\u2030", 3],
  ["\u052A\u20A7", 3.1],
  ["\u2194", 3.25],
  ["\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469\u246A\u246B\u246C\u246D\u246E\u246F\u2470\u2471\u2472\u2473\u24EA\u24EB\u24EC\u24ED\u24EE\u24EF\u24F0\u24F1\u24F2\u24F3\u24F4\u24FF", 3.35],
  ["\u2167", 3.45],
  ["\u2031", 3.6]
]);
function textDisplayUnits(text, fontName) {
  const costs = usesNarrowFont(fontName) ? NARROW_FONT_COSTS : WIDE_FONT_COSTS;
  let units = 0;
  for (const char of text) {
    const cp = char.codePointAt(0);
    if (isZeroWidthCodePoint(cp)) continue;
    if (isEmojiCodePoint(cp)) {
      units += 3.4;
      continue;
    }
    if (isCjkWideCodePoint(cp)) {
      units += 2.2;
      continue;
    }
    const cost = costs[char] ?? WIDE_UNICODE_COSTS[char];
    if (cost !== void 0) {
      units += cost;
      continue;
    }
    units += cp < 592 ? 1.3 : 2;
  }
  return units;
}
function fontScale(font) {
  let scale = Math.max(1, (font?.size ?? DEFAULT_FONT_SIZE) / DEFAULT_FONT_SIZE);
  if (font?.bold) scale *= BOLD_WIDTH_FACTOR;
  if (font?.italic) scale *= ITALIC_WIDTH_FACTOR;
  return scale;
}
function alignmentIndentChars(alignment) {
  const indent = alignment?.indent ?? 0;
  if (!Number.isFinite(indent) || indent <= 0) return 0;
  const sides = alignment?.horizontal === "center" || alignment?.horizontal === "distributed" ? 2 : 1;
  return indent * sides * INDENT_WIDTH;
}
function autofitUnitsWidth(units, scale = 1, extraChars = 0) {
  const width = (AUTOFIT_FACTOR * (units * LONG_TEXT_FACTOR + 1) + 2) * scale + extraChars;
  return Math.min(MAX_AUTOFIT_WIDTH, Math.max(DEFAULT_COLUMN_WIDTH, width));
}
function autofitWidth(text, font, indentChars = 0) {
  return autofitUnitsWidth(textDisplayUnits(text, font?.name), fontScale(font), indentChars);
}
function autofitFileWidth(text, font, indentChars = 0) {
  return Math.min(
    MAX_AUTOFIT_WIDTH,
    autofitWidth(text, font, indentChars) + COLUMN_WIDTH_PADDING
  );
}
function autofitFileUnitsWidth(units, indentChars = 0) {
  return Math.min(
    MAX_AUTOFIT_WIDTH,
    autofitUnitsWidth(units, 1, indentChars) + COLUMN_WIDTH_PADDING
  );
}
function richTextUnits(rich, baseFont) {
  let units = 0;
  for (const run of rich.runs) {
    const font = run.style?.font ?? baseFont;
    units += textDisplayUnits(run.text, font?.name) * fontScale(font);
  }
  return units;
}
function richTextLineUnits(rich, baseFont) {
  const lines = [0];
  for (const run of rich.runs) {
    const font = run.style?.font ?? baseFont;
    const parts = run.text.split("\n");
    for (let index = 0; index < parts.length; index++) {
      if (index > 0) lines.push(0);
      lines[lines.length - 1] += textDisplayUnits(parts[index], font?.name) * fontScale(font);
    }
  }
  return lines;
}
var DATE_TOKEN_AT = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\.0+)/i;
var DATE_CHECK = /(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p)/i;
var DATE_SAMPLES = {
  yyyy: "2024",
  yy: "24",
  mmmmm: "W",
  mmmm: "pa\u017Adziernik",
  mmm: "Sep",
  mm: "12",
  m: "12",
  dddd: "poniedzia\u0142ek",
  ddd: "Wed",
  dd: "30",
  d: "30",
  hh: "23",
  h: "23",
  ss: "59",
  s: "59",
  "am/pm": "PM",
  "a/p": "P"
};
function formatSections(format) {
  const sections = [];
  let section = "";
  let quoted = false;
  for (const char of format) {
    if (char === '"') quoted = !quoted;
    if (char === ";" && !quoted) {
      sections.push(section);
      section = "";
    } else {
      section += char;
    }
  }
  sections.push(section);
  return sections;
}
function withoutQuotedText(format) {
  return format.replace(/"[^"]*"/g, "");
}
function dateFormatText(section) {
  if (!DATE_CHECK.test(withoutQuotedText(section))) return null;
  let result = "";
  let i = 0;
  while (i < section.length) {
    const char = section[i];
    if (char === '"') {
      const end = section.indexOf('"', i + 1);
      if (end < 0) {
        result += section.slice(i + 1);
        break;
      }
      result += section.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (char === "\\" && i + 1 < section.length) {
      result += section[i + 1];
      i += 2;
      continue;
    }
    if (char === "[") {
      const end = section.indexOf("]", i + 1);
      i = end < 0 ? section.length : end + 1;
      continue;
    }
    const token = DATE_TOKEN_AT.exec(section.slice(i));
    if (token) {
      const key = token[1].toLowerCase();
      if (key.startsWith(".0")) {
        result += "." + "0".repeat(token[1].length - 1);
      } else {
        result += DATE_SAMPLES[key] ?? token[1];
      }
      i += token[1].length;
      continue;
    }
    result += char;
    i++;
  }
  return result;
}
function formatLiteralText(section) {
  let result = "";
  let i = 0;
  while (i < section.length) {
    const char = section[i];
    if (char === '"') {
      const end = section.indexOf('"', i + 1);
      if (end < 0) {
        result += section.slice(i + 1);
        break;
      }
      result += section.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (char === "\\" && i + 1 < section.length) {
      result += section[i + 1];
      i += 2;
      continue;
    }
    if (char === "[") {
      const end = section.indexOf("]", i + 1);
      if (end < 0) break;
      const content = section.slice(i + 1, end);
      if (content.startsWith("$")) {
        result += content.slice(1).replace(/-?\d+$/, "").trim();
      }
      i = end + 1;
      continue;
    }
    if ("0#?.,@".includes(char)) {
      i++;
      continue;
    }
    if ("Ee+-".includes(char)) {
      i++;
      continue;
    }
    result += char;
    i++;
  }
  return result;
}
function expandExponential(text) {
  const match = /^(-?)(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/.exec(text);
  if (!match) return text;
  const sign = match[1];
  const integer = match[2];
  const fraction = match[3] ?? "";
  const digits = integer + fraction;
  const point = integer.length + Number(match[4]);
  if (point <= 0) return `${sign}0.${"0".repeat(-point)}${digits}`;
  if (point >= digits.length) return `${sign}${digits}${"0".repeat(point - digits.length)}`;
  return `${sign}${digits.slice(0, point)}.${digits.slice(point)}`;
}
function scientificText(value) {
  const [mantissaRaw, exponentRaw] = Math.abs(value).toExponential(5).split("e");
  const mantissa = mantissaRaw.replace(/\.?0+$/, "");
  const exponent = Number(exponentRaw);
  return `${mantissa}E${exponent < 0 ? "-" : "+"}${String(Math.abs(exponent)).padStart(2, "0")}`;
}
function generalNumberDisplay(value) {
  if (!Number.isFinite(value)) {
    if (Number.isNaN(value)) return "NaN";
    return value < 0 ? "-\u221E" : "\u221E";
  }
  if (value === 0) return "0";
  const sign = value < 0 ? "-" : "";
  const abs = Math.abs(value);
  if (Number.isInteger(abs)) {
    const text = String(abs);
    if (!/[eE]/.test(text) && text.length <= 11) return sign + text;
    return sign + scientificText(abs);
  }
  if (abs >= 1e-9 && abs < 1e11) {
    const whole = Math.trunc(abs);
    const wholeLength = String(whole).length;
    const decimals = Math.max(0, 10 - wholeLength);
    const rounded = Number(abs.toFixed(decimals));
    let fixed = rounded.toFixed(decimals);
    if (fixed.includes(".")) fixed = fixed.replace(/0+$/, "").replace(/\.$/, "");
    const exact = rounded === abs;
    const room = whole > 0 ? decimals + wholeLength : decimals + Math.floor(Math.log10(abs)) + 1;
    if (fixed !== "0" && (exact || room >= 6)) return sign + fixed;
  }
  return sign + scientificText(abs);
}
function generalNumberText(value, expandLarge) {
  if (!Number.isFinite(value)) return String(value);
  let text = String(value);
  if (/[eE]/.test(text) && (expandLarge || Math.abs(value) < 1e-6)) {
    text = expandExponential(text);
  }
  if (text.endsWith(".0")) text = text.slice(0, -2);
  return text;
}
function generalValueText(value, expandLarge = false) {
  if (value === null || value === void 0) return "";
  if (typeof value === "boolean") return "PRAWDA";
  if (value instanceof Date) {
    return dateHasTime(value) ? "12/31/2024 23:45" : "12/31/2024";
  }
  if (typeof value === "number") {
    return expandLarge ? generalNumberText(value, true) : generalNumberDisplay(value);
  }
  if (typeof value === "bigint") return value.toString();
  if (value instanceof ExcelError) return value.code;
  return String(value);
}
function estimateCellDisplayText(value, numFmt) {
  let raw = generalValueText(value);
  if (!numFmt || /^general$/i.test(numFmt.trim())) return raw;
  const sections = formatSections(numFmt);
  const negative = typeof value === "bigint" ? value < 0n : typeof value === "number" && value < 0;
  const secondSection = sections[1] ?? "";
  const useNegativeSection = negative && secondSection.trim() !== "";
  const section = useNegativeSection ? secondSection : sections[0] ?? "";
  const displayValue = useNegativeSection ? typeof value === "bigint" ? -value : Math.abs(value) : value;
  if (typeof displayValue !== "number" && typeof displayValue !== "bigint" && !(displayValue instanceof Date)) {
    return raw + formatLiteralText(section);
  }
  raw = generalValueText(displayValue, true);
  const dateText = dateFormatText(section);
  if (dateText !== null) {
    return textDisplayUnits(dateText) >= textDisplayUnits(raw) ? dateText : raw + dateText;
  }
  let result = raw + formatLiteralText(section);
  const clean = withoutQuotedText(section).replace(/\[[^\]]*\]/g, "");
  const placeholder = /[0#?][0#?,.]*/.exec(clean);
  if (placeholder) {
    const [integerPart, fractionPart = ""] = placeholder[0].split(".");
    const minIntegerDigits = (integerPart.match(/0/g) ?? []).length;
    const integerMatch = /^-?\d+/.exec(raw);
    const rawIntegerDigits = integerMatch ? integerMatch[0].replace("-", "").length : 0;
    result += "0".repeat(Math.max(0, minIntegerDigits - rawIntegerDigits));
    const decimals = (fractionPart.match(/[0#?]/g) ?? []).length;
    const rawFractionDigits = raw.includes(".") ? raw.split(".")[1].length : 0;
    const extraDecimals = Math.max(0, decimals - rawFractionDigits);
    if (extraDecimals > 0) {
      result += (raw.includes(".") ? "" : ".") + "0".repeat(extraDecimals);
    }
    if (integerPart.includes(",")) {
      const digits = (raw.match(/\d/g) ?? []).length;
      result += ",".repeat(Math.max(0, Math.floor((digits - 1) / 3)));
    }
  }
  const exponent = /[eE][+-]([0#]+)/.exec(clean);
  if (exponent) result += "E+" + "0".repeat(Math.max(2, exponent[1].length));
  const percent = (clean.match(/%/g) ?? []).length;
  if (percent > 0) result += "00".repeat(percent) + "%".repeat(percent);
  return result;
}
function applyHeaderWidths(colWidths, headers, columnCount, options) {
  const font = options?.bold ? { bold: true } : null;
  for (let i = 0; i < columnCount; i++) {
    let width = autofitFileWidth(headers[i] ?? "", font);
    if (options?.filterButton) {
      width = Math.min(MAX_AUTOFIT_WIDTH, width + AUTOFILTER_BUTTON_WIDTH);
    }
    if (colWidths[i] < width) colWidths[i] = width;
  }
}
function richTextFileWidth(rich, baseFont, alignment) {
  const indent = alignmentIndentChars(alignment);
  if (alignment?.wrapText && rich.runs.some((run) => run.text.includes("\n"))) {
    let best = 0;
    for (const units of richTextLineUnits(rich, baseFont)) {
      best = Math.max(best, autofitFileUnitsWidth(units, indent));
    }
    return best;
  }
  return autofitFileUnitsWidth(richTextUnits(rich, baseFont), indent);
}
function cellAutofitFileWidth(cell) {
  if (isRichText(cell)) return richTextFileWidth(cell);
  const style = getCellStyle(cell);
  const value = unwrapCell(cell);
  if (value === null || value === void 0) return -1;
  const indent = alignmentIndentChars(style?.alignment);
  if (isRichText(value)) return richTextFileWidth(value, style?.font, style?.alignment);
  const text = estimateCellDisplayText(value, style?.numFmt);
  if (style?.alignment?.wrapText && text.includes("\n")) {
    let best = 0;
    for (const line of text.split("\n")) {
      best = Math.max(best, autofitFileWidth(line, style?.font, indent));
    }
    return best;
  }
  return autofitFileWidth(text, style?.font, indent);
}
function updateColWidthsFromRows(colWidths, rows, sampleLimit = AUTOFIT_SAMPLE_ROWS) {
  for (let r = 0; r < Math.min(rows.length, sampleLimit); r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const raw = row[c];
      if (raw === null || raw === void 0) continue;
      const width = cellAutofitFileWidth(raw);
      if (width >= 0 && colWidths[c] < width) colWidths[c] = width;
    }
  }
}
function defaultColWidth(width, floor = false) {
  const value = width > 0 ? width : DEFAULT_COLUMN_WIDTH + COLUMN_WIDTH_PADDING;
  return floor ? Math.floor(value) : value;
}
function assertColumnWidth(width, index) {
  if (typeof width !== "number" || !Number.isFinite(width)) {
    throw new TypeError(`columnWidths[${index}] must be a finite number`);
  }
  if (width < 0.5 || width > 255) {
    throw new RangeError(`columnWidths[${index}] must be between 0.5 and 255, got ${width}`);
  }
  return width;
}
function formatColumnWidth(width) {
  return String(Math.round(width * 100) / 100);
}
function cellRefAttr(col, prevCol, colRef, rowNum) {
  return col === prevCol + 1 ? "" : ` r="${colRef}${rowNum}"`;
}

// src/writers/core/xlsxParts.ts
function escapeXlsxAttribute(value) {
  return escapeXmlText(value);
}
function formatSheetNameForFormula(name) {
  if (/[\s\-+=()!@#$%^&]/.test(name) || /^[0-9]/.test(name)) {
    return `'${name.replace(/'/g, "''")}'`;
  }
  return name;
}
var BORDER_LINE_STYLES = /* @__PURE__ */ new Set([
  "thin",
  "medium",
  "dashed",
  "dotted",
  "thick",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot"
]);
var PATTERN_TYPES = /* @__PURE__ */ new Set([
  "none",
  "solid",
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625"
]);
var HORIZONTAL_ALIGNMENTS = /* @__PURE__ */ new Set([
  "general",
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed"
]);
var VERTICAL_ALIGNMENTS = /* @__PURE__ */ new Set(["top", "center", "bottom", "justify", "distributed"]);
function colorKey(color) {
  if (typeof color === "string") {
    let value = color.trim().replace(/^#/, "").toUpperCase();
    if (value.length === 6) value = `FF${value}`;
    if (!/^[0-9A-F]{8}$/.test(value)) {
      throw new Error(`Invalid color '${color}'. Use RRGGBB, AARRGGBB or '#' + hex.`);
    }
    return `rgb:${value}`;
  }
  if ("indexed" in color) return `indexed:${color.indexed}`;
  const tint = color.tint ?? 0;
  return `theme:${color.theme}:${tint}`;
}
function colorXml(color) {
  if (color === void 0) return "";
  if (typeof color === "string") {
    const key = colorKey(color);
    return `<color rgb="${key.slice(4)}"/>`;
  }
  if ("indexed" in color) return `<color indexed="${color.indexed}"/>`;
  const tint = color.tint;
  return `<color theme="${color.theme}"${tint !== void 0 ? ` tint="${tint}"` : ""}/>`;
}
function normalizeFont(font) {
  if (!font) return null;
  const out = {};
  if (font.name !== void 0) out.name = font.name;
  if (font.size !== void 0) {
    if (!Number.isFinite(font.size) || font.size <= 0 || font.size > 409) {
      throw new RangeError(`font.size must be in (0, 409], got ${font.size}`);
    }
    out.size = font.size;
  }
  if (font.bold) out.bold = true;
  if (font.italic) out.italic = true;
  if (font.strike) out.strike = true;
  if (font.outline) out.outline = true;
  if (font.shadow) out.shadow = true;
  if (font.condense) out.condense = true;
  if (font.extend) out.extend = true;
  if (font.underline === true) out.underline = "single";
  else if (typeof font.underline === "string") out.underline = font.underline;
  if (font.color !== void 0) out.color = font.color;
  if (font.family !== void 0) out.family = font.family;
  if (font.scheme !== void 0) out.scheme = font.scheme;
  if (font.vertAlign !== void 0 && font.vertAlign !== "baseline") out.vertAlign = font.vertAlign;
  if (font.charset !== void 0) out.charset = font.charset;
  return Object.keys(out).length > 0 ? out : null;
}
function normalizeFill(fill) {
  if (!fill) return null;
  const out = {};
  if (fill.patternType !== void 0) {
    if (!PATTERN_TYPES.has(fill.patternType)) throw new Error(`Unknown fill pattern '${fill.patternType}'`);
    out.patternType = fill.patternType;
  }
  if (fill.fgColor !== void 0) out.fgColor = fill.fgColor;
  if (fill.bgColor !== void 0) out.bgColor = fill.bgColor;
  if (Object.keys(out).length === 0) return null;
  if (out.patternType === void 0) out.patternType = "solid";
  return out;
}
function normalizeBorderSide(side) {
  if (!side) return null;
  if (side.style !== void 0 && !BORDER_LINE_STYLES.has(side.style)) {
    throw new Error(`Unknown border style '${side.style}'`);
  }
  const out = {};
  if (side.style !== void 0) out.style = side.style;
  if (side.color !== void 0) out.color = side.color;
  return Object.keys(out).length > 0 ? out : null;
}
function normalizeBorder(border) {
  if (!border) return null;
  const out = {};
  for (const name of ["top", "bottom", "left", "right", "diagonal"]) {
    const side = normalizeBorderSide(border[name]);
    if (side) out[name] = side;
  }
  if (border.diagonalUp) out.diagonalUp = true;
  if (border.diagonalDown) out.diagonalDown = true;
  if (border.outline === false) out.outline = false;
  return Object.keys(out).length > 0 ? out : null;
}
function normalizeAlignment(alignment) {
  if (!alignment) return null;
  const out = {};
  if (alignment.horizontal !== void 0) {
    if (!HORIZONTAL_ALIGNMENTS.has(alignment.horizontal)) {
      throw new Error(`Unknown horizontal alignment '${alignment.horizontal}'`);
    }
    out.horizontal = alignment.horizontal;
  }
  if (alignment.vertical !== void 0) {
    if (!VERTICAL_ALIGNMENTS.has(alignment.vertical)) {
      throw new Error(`Unknown vertical alignment '${alignment.vertical}'`);
    }
    out.vertical = alignment.vertical;
  }
  if (alignment.wrapText) out.wrapText = true;
  if (alignment.shrinkToFit) out.shrinkToFit = true;
  if (alignment.indent !== void 0) {
    if (!Number.isInteger(alignment.indent) || alignment.indent < 0 || alignment.indent > 250) {
      throw new RangeError(`alignment.indent must be 0..250, got ${alignment.indent}`);
    }
    out.indent = alignment.indent;
  }
  if (alignment.textRotation !== void 0) {
    const rotation = alignment.textRotation;
    if (!Number.isInteger(rotation) || !(rotation >= 0 && rotation <= 180 || rotation === 255)) {
      throw new RangeError(`alignment.textRotation must be 0..180 or 255, got ${rotation}`);
    }
    out.textRotation = rotation;
  }
  if (alignment.justifyLastLine) out.justifyLastLine = true;
  if (alignment.readingOrder !== void 0) out.readingOrder = alignment.readingOrder;
  return Object.keys(out).length > 0 ? out : null;
}
function normalizeProtection(protection) {
  if (!protection) return null;
  const out = {};
  if (protection.locked === false) out.locked = false;
  if (protection.hidden) out.hidden = true;
  return Object.keys(out).length > 0 ? out : null;
}
function normalizeStyle(style) {
  if (!style) return null;
  const out = {};
  if (style.numFmt !== void 0 && style.numFmt !== "") out.numFmt = style.numFmt;
  const font = normalizeFont(style.font);
  const fill = normalizeFill(style.fill);
  const border = normalizeBorder(style.border);
  const alignment = normalizeAlignment(style.alignment);
  const protection = normalizeProtection(style.protection);
  if (font) out.font = font;
  if (fill) out.fill = fill;
  if (border) out.border = border;
  if (alignment) out.alignment = alignment;
  if (protection) out.protection = protection;
  return Object.keys(out).length > 0 ? out : null;
}
var BASE_FONT_NORMAL = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
var BASE_FONT_BOLD = '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
var BASE_FILL_NONE = '<fill><patternFill patternType="none"/></fill>';
var BASE_FILL_GRAY = '<fill><patternFill patternType="gray125"/></fill>';
var BASE_FILL_SOLID = '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>';
var BASE_BORDER = "<border><left/><right/><top/><bottom/><diagonal/></border>";
var BASE_XFS = [
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
  '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
  '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
];
function fontXml(font) {
  let xml = "<font>";
  if (font.name !== void 0) xml += `<name val="${escapeXlsxAttribute(font.name)}"/>`;
  if (font.charset !== void 0) xml += `<charset val="${font.charset}"/>`;
  if (font.family !== void 0) xml += `<family val="${font.family}"/>`;
  if (font.bold) xml += "<b/>";
  if (font.italic) xml += "<i/>";
  if (font.strike) xml += "<strike/>";
  if (font.outline) xml += "<outline/>";
  if (font.shadow) xml += "<shadow/>";
  if (font.condense) xml += "<condense/>";
  if (font.extend) xml += "<extend/>";
  if (font.color !== void 0) xml += colorXml(font.color);
  if (font.size !== void 0) xml += `<sz val="${font.size}"/>`;
  if (font.underline !== void 0) {
    xml += font.underline === "single" ? "<u/>" : `<u val="${font.underline}"/>`;
  }
  if (font.vertAlign !== void 0) xml += `<vertAlign val="${font.vertAlign}"/>`;
  if (font.scheme !== void 0) xml += `<scheme val="${font.scheme}"/>`;
  return `${xml}</font>`;
}
function fillXml(fill) {
  const pattern = fill.patternType ?? "solid";
  let xml = `<fill><patternFill patternType="${pattern}">`;
  if (fill.fgColor !== void 0) xml += colorXml(fill.fgColor).replace("<color", "<fgColor");
  if (fill.bgColor !== void 0) xml += colorXml(fill.bgColor).replace("<color", "<bgColor");
  return `${xml}</patternFill></fill>`;
}
function borderSideXml(tag, side) {
  if (!side || side.style === void 0 && side.color === void 0) return `<${tag}/>`;
  let xml = `<${tag}`;
  if (side.style !== void 0) xml += ` style="${side.style}"`;
  xml += ">";
  if (side.color !== void 0) xml += colorXml(side.color);
  return `${xml}</${tag}>`;
}
function borderXml(border) {
  let xml = "<border";
  if (border.diagonalUp) xml += ' diagonalUp="1"';
  if (border.diagonalDown) xml += ' diagonalDown="1"';
  if (border.outline === false) xml += ' outline="0"';
  xml += ">";
  xml += borderSideXml("left", border.left);
  xml += borderSideXml("right", border.right);
  xml += borderSideXml("top", border.top);
  xml += borderSideXml("bottom", border.bottom);
  xml += borderSideXml("diagonal", border.diagonal);
  return `${xml}</border>`;
}
function alignmentXml(alignment) {
  let xml = "<alignment";
  if (alignment.horizontal !== void 0) xml += ` horizontal="${alignment.horizontal}"`;
  if (alignment.vertical !== void 0) xml += ` vertical="${alignment.vertical}"`;
  if (alignment.wrapText) xml += ' wrapText="1"';
  if (alignment.shrinkToFit) xml += ' shrinkToFit="1"';
  if (alignment.indent !== void 0) xml += ` indent="${alignment.indent}"`;
  if (alignment.textRotation !== void 0) xml += ` textRotation="${alignment.textRotation}"`;
  if (alignment.justifyLastLine) xml += ' justifyLastLine="1"';
  if (alignment.readingOrder !== void 0) xml += ` readingOrder="${alignment.readingOrder}"`;
  return `${xml}/>`;
}
function protectionXml(protection) {
  let xml = "<protection";
  if (protection.locked === false) xml += ' locked="0"';
  if (protection.hidden) xml += ' hidden="1"';
  return `${xml}/>`;
}
function dxfXml(style, numFmtId) {
  let xml = "<dxf>";
  if (style.font) xml += fontXml(style.font);
  if (style.numFmt !== void 0 && numFmtId !== null) {
    xml += `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXlsxAttribute(style.numFmt)}"/>`;
  }
  if (style.fill) {
    const color = style.fill.bgColor ?? style.fill.fgColor;
    xml += "<fill><patternFill>";
    if (color !== void 0) xml += colorXml(color).replace("<color", "<bgColor");
    xml += "</patternFill></fill>";
  }
  if (style.alignment) xml += alignmentXml(style.alignment);
  if (style.border) xml += borderXml(style.border);
  if (style.protection) xml += protectionXml(style.protection);
  xml += "</dxf>";
  return xml;
}
var XlsxStyleRegistry = class {
  constructor() {
    this.numFmts = /* @__PURE__ */ new Map();
    this.fontIds = /* @__PURE__ */ new Map();
    this.fillIds = /* @__PURE__ */ new Map();
    this.borderIds = /* @__PURE__ */ new Map();
    this.xfIds = /* @__PURE__ */ new Map();
    this.fonts = [BASE_FONT_NORMAL, BASE_FONT_BOLD];
    this.fills = [BASE_FILL_NONE, BASE_FILL_GRAY, BASE_FILL_SOLID];
    this.borders = [BASE_BORDER];
    this.xfs = [...BASE_XFS];
    this.styles = [null, null, null, null, null, null];
    this.dxfIds = /* @__PURE__ */ new Map();
    this.dxfs = [];
    this.nextNumFmtId = 165;
  }
  numFmtIdFor(format) {
    let id = this.numFmts.get(format);
    if (id === void 0) {
      id = this.nextNumFmtId++;
      this.numFmts.set(format, id);
    }
    return id;
  }
  /** Cell XF index for a style; 0 for null/empty styles. */
  register(style) {
    const normalized = normalizeStyle(style);
    if (normalized === null) return 0;
    const key = JSON.stringify(normalized);
    const existing = this.xfIds.get(key);
    if (existing !== void 0) return existing;
    let numFmtId = 0;
    if (normalized.numFmt !== void 0) {
      numFmtId = this.numFmtIdFor(normalized.numFmt);
    }
    let fontId = 0;
    if (normalized.font) {
      if (normalized.font.bold && Object.keys(normalized.font).length === 1) {
        fontId = 1;
      } else {
        const componentKey = JSON.stringify(normalized.font);
        let id = this.fontIds.get(componentKey);
        if (id === void 0) {
          id = this.fonts.length;
          this.fontIds.set(componentKey, id);
          this.fonts.push(fontXml(normalized.font));
        }
        fontId = id;
      }
    }
    let fillId = 0;
    if (normalized.fill) {
      const componentKey = JSON.stringify(normalized.fill);
      let id = this.fillIds.get(componentKey);
      if (id === void 0) {
        id = this.fills.length;
        this.fillIds.set(componentKey, id);
        this.fills.push(fillXml(normalized.fill));
      }
      fillId = id;
    }
    let borderId = 0;
    if (normalized.border) {
      const componentKey = JSON.stringify(normalized.border);
      let id = this.borderIds.get(componentKey);
      if (id === void 0) {
        id = this.borders.length;
        this.borderIds.set(componentKey, id);
        this.borders.push(borderXml(normalized.border));
      }
      borderId = id;
    }
    let xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
    if (numFmtId !== 0) xf += ' applyNumberFormat="1"';
    if (fontId !== 0) xf += ' applyFont="1"';
    if (fillId !== 0) xf += ' applyFill="1"';
    if (borderId !== 0) xf += ' applyBorder="1"';
    if (normalized.alignment) xf += ' applyAlignment="1"';
    if (normalized.protection) xf += ' applyProtection="1"';
    const children = (normalized.alignment ? alignmentXml(normalized.alignment) : "") + (normalized.protection ? protectionXml(normalized.protection) : "");
    if (children) {
      xf += `>${children}</xf>`;
    } else {
      xf += "/>";
    }
    const index = this.xfs.length;
    this.xfIds.set(key, index);
    this.xfs.push(xf);
    this.styles.push(normalized);
    return index;
  }
  /** Differential style id for conditional formatting; -1 for an empty style. */
  registerDxf(style) {
    const normalized = normalizeStyle(style);
    if (normalized === null) return -1;
    const key = JSON.stringify(normalized);
    const existing = this.dxfIds.get(key);
    if (existing !== void 0) return existing;
    const numFmtId = normalized.numFmt !== void 0 ? this.numFmtIdFor(normalized.numFmt) : null;
    const index = this.dxfs.length;
    this.dxfIds.set(key, index);
    this.dxfs.push(dxfXml(normalized, numFmtId));
    return index;
  }
  /** XF index for a plain number format (legacy FormattedCell path). */
  registerFormat(format) {
    return this.register({ numFmt: format });
  }
  toXml() {
    let numFmts = "";
    for (const [format, id] of this.numFmts) {
      numFmts += `<numFmt numFmtId="${id}" formatCode="${escapeXlsxAttribute(format)}"/>`;
    }
    const numFmtsSection = this.numFmts.size > 0 ? `<numFmts count="${this.numFmts.size}">${numFmts}</numFmts>` : "";
    const dxfsSection = this.dxfs.length > 0 ? `<dxfs count="${this.dxfs.length}">${this.dxfs.join("\n")}</dxfs>` : '<dxfs count="0"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${numFmtsSection}
<fonts count="${this.fonts.length}">
${this.fonts.join("\n")}
</fonts>
<fills count="${this.fills.length}">
${this.fills.join("\n")}
</fills>
<borders count="${this.borders.length}">
${this.borders.join("\n")}
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="${this.xfs.length}">
${this.xfs.join("\n")}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
${dxfsSection}
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }
};
function sharedStringKey(value) {
  if (typeof value === "string") return `s:${value}`;
  return `r:${JSON.stringify(value.runs)}`;
}
function richTextXml(rich) {
  let xml = "";
  for (const run of rich.runs) {
    xml += "<r>";
    const font = normalizeFont(run.style?.font);
    if (font) {
      xml += fontXml(font).replace("<font>", "<rPr>").replace("</font>", "</rPr>");
    }
    xml += `<t xml:space="preserve">${escapeXlsxText(run.text)}</t></r>`;
  }
  return xml;
}
function sharedStringItem(value) {
  if (typeof value !== "string") return `<si>${richTextXml(value)}</si>`;
  const clean = escapeXlsxText(value);
  if (clean.length > 0 && (clean[0] === " " || clean[clean.length - 1] === " " || /[\t\n\r]/.test(clean))) {
    return `<si><t xml:space="preserve">${clean}</t></si>`;
  }
  return `<si><t>${clean}</t></si>`;
}
function buildXlsxSharedStrings(strings, countAll) {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${countAll}" uniqueCount="${strings.length}">`;
  for (const value of strings) xml += sharedStringItem(value);
  xml += "</sst>";
  return xml;
}
function buildXlsxWorkbook(sheets, autofilterOn) {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>`;
  for (const sheet of sheets) {
    const state = sheet.hidden ? ' state="hidden"' : "";
    xml += `<sheet name="${escapeXlsxAttribute(sheet.name)}" sheetId="${sheet.sheetId}"${state} r:id="${sheet.rId}"/>`;
  }
  xml += "</sheets>";
  if (autofilterOn) {
    xml += "<definedNames>";
    for (const sheet of sheets) {
      if (!sheet.filterHeaderRange) continue;
      const escapedRange = escapeXlsxText(sheet.filterHeaderRange);
      xml += `<definedName name="_xlnm._FilterDatabase" localSheetId="${sheet.sheetId - 1}" hidden="1">${escapedRange}</definedName>`;
    }
    xml += "</definedNames>";
  }
  xml += '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>';
  return xml;
}
function buildXlsxContentTypes(sheets, comments) {
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${comments && comments.length > 0 ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : ""}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
  for (const sheet of sheets) {
    xml += `<Override PartName="/${sheet.pathInArchive}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  }
  for (const part of comments ?? []) {
    xml += `<Override PartName="/${part.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
  }
  xml += "</Types>";
  return xml;
}
function buildXlsxRels(sheets) {
  const global = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  let workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
  for (const sheet of sheets) {
    workbook += `<Relationship Id="${sheet.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheet.nameInArchive}"/>`;
  }
  let nextId = sheets.length + 1;
  workbook += `<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  workbook += `<Relationship Id="rId${nextId++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
  workbook += "</Relationships>";
  return { global, workbook };
}
function writeXlsxSheetHead(sink, options) {
  const { dimensionRef, columnCount, isFirstSheet, hasHeaders, colWidths } = options;
  sink.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  sink.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
  sink.writeString(`<dimension ref="${dimensionRef}"/>`);
  if (hasHeaders) {
    sink.writeString(isFirstSheet ? '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' : '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
  } else {
    sink.writeString(isFirstSheet ? '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' : '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
  }
  sink.writeString("<cols>");
  for (let i = 0; i < columnCount; i++) {
    const width = defaultColWidth(colWidths[i]);
    sink.writeString(`<col min="${i + 1}" max="${i + 1}" width="${formatColumnWidth(width)}" bestFit="1" customWidth="1" />`);
  }
  sink.writeString("</cols><sheetData>");
}
function writeXlsxSheetTail(sink, options) {
  const {
    doAutofilter,
    columnCount,
    lastColumnLetter,
    rowNum,
    sheetName,
    merges,
    hyperlinks,
    dataValidations,
    conditionalFormats,
    legacyDrawingRId,
    cfPriorityStart = 1,
    registerDxf
  } = options;
  sink.writeString("</sheetData>");
  let filterRange = null;
  if (doAutofilter && columnCount > 0) {
    const ref = `A1:${lastColumnLetter}${rowNum}`;
    sink.writeString(`<autoFilter ref="${ref}"/>`);
    filterRange = `${formatSheetNameForFormula(sheetName)}!$A$1:$${lastColumnLetter}$${rowNum}`;
  }
  if (merges && merges.length > 0) {
    sink.writeString(`<mergeCells count="${merges.length}">`);
    for (const ref of merges) sink.writeString(`<mergeCell ref="${escapeXmlText(ref)}"/>`);
    sink.writeString("</mergeCells>");
  }
  if (conditionalFormats && conditionalFormats.length > 0) {
    writeConditionalFormatting(sink, conditionalFormats, cfPriorityStart, registerDxf);
  }
  if (dataValidations && dataValidations.length > 0) {
    writeDataValidations(sink, dataValidations);
  }
  if (hyperlinks && hyperlinks.length > 0) {
    writeHyperlinks(sink, hyperlinks);
  }
  if (legacyDrawingRId) {
    sink.writeString(`<legacyDrawing r:id="${legacyDrawingRId}"/>`);
  }
  sink.writeString("</worksheet>");
  return filterRange;
}
function columnLetterFromRef(ref) {
  const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref.trim());
  if (!match) throw new Error(`Invalid cell reference '${ref}'`);
  let col = 0;
  for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col: col - 1, row: parseInt(match[2], 10) - 1 };
}
function buildXlsxComments(comments) {
  const authors = [];
  const authorIds = /* @__PURE__ */ new Map();
  for (const comment of comments) {
    const author = comment.author ?? "";
    if (!authorIds.has(author)) {
      authorIds.set(author, authors.length);
      authors.push(author);
    }
  }
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml += '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>';
  for (const author of authors) xml += `<author>${escapeXlsxText(author)}</author>`;
  xml += "</authors><commentList>";
  for (const comment of comments) {
    const text = escapeXlsxText(comment.text);
    const space = text.length > 0 && (text[0] === " " || text[text.length - 1] === " " || /[\t\n\r]/.test(text)) ? ' xml:space="preserve"' : "";
    xml += `<comment ref="${escapeXmlText(comment.ref)}" authorId="${authorIds.get(comment.author ?? "")}"><text><t${space}>${text}</t></text></comment>`;
  }
  xml += "</commentList></comments>";
  return xml;
}
function buildXlsxCommentsVml(comments) {
  const V = "urn:schemas-microsoft-com:vml";
  const O = "urn:schemas-microsoft-com:office:office";
  const X = "urn:schemas-microsoft-com:office:excel";
  let xml = `<xml xmlns:v="${V}" xmlns:o="${O}" xmlns:x="${X}">`;
  xml += `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`;
  let nextId = 1025;
  for (const comment of comments) {
    const { col, row } = columnLetterFromRef(comment.ref);
    xml += `<v:shape id="_x0000_s${nextId++}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
  }
  xml += "</xml>";
  return xml;
}
function buildXlsxSheetRels(rels) {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  for (const rel of rels) {
    xml += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXlsxAttribute(rel.target)}"`;
    if (rel.external) xml += ' TargetMode="External"';
    xml += "/>";
  }
  xml += "</Relationships>";
  return xml;
}
var XLSX_REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
var XLSX_REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
var XLSX_REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
function writeXlsxSharedStringCell(sink, sst, value, ref, styleOverride) {
  const key = sharedStringKey(value);
  let index = sst.index.get(key);
  if (index === void 0) {
    index = sst.values.length;
    sst.values.push(value);
    sst.index.set(key, index);
  }
  const styleAttr = styleOverride !== void 0 ? ` s="${styleOverride}"` : "";
  sink.writeString(`<c${ref} t="s"${styleAttr}><v>${index}</v></c>`);
}
function writeXlsxInlineStringCell(sink, value, ref, styleOverride) {
  const styleAttr = styleOverride !== void 0 ? ` s="${styleOverride}"` : "";
  if (typeof value !== "string") {
    sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is>${richTextXml(value)}</is></c>`);
    return;
  }
  const clean = escapeXlsxText(value);
  if (clean.length > 0 && (clean[0] === " " || clean[clean.length - 1] === " " || /[\t\n\r]/.test(clean))) {
    sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is><t xml:space="preserve">${clean}</t></is></c>`);
  } else {
    sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is><t>${clean}</t></is></c>`);
  }
}
var DV_TYPES = {
  whole: "whole",
  decimal: "decimal",
  list: "list",
  date: "date",
  time: "time",
  textLength: "textLength",
  custom: "custom"
};
var DV_OPERATORS = {
  between: "between",
  notBetween: "notBetween",
  equal: "equal",
  notEqual: "notEqual",
  lessThan: "lessThan",
  lessThanOrEqual: "lessThanOrEqual",
  greaterThan: "greaterThan",
  greaterThanOrEqual: "greaterThanOrEqual"
};
function writeDataValidations(sink, validations) {
  if (validations.length === 0) return;
  sink.writeString(`<dataValidations count="${validations.length}">`);
  for (const rule of validations) {
    let attrs = ` type="${DV_TYPES[rule.type]}" sqref="${escapeXmlText(rule.sqref)}"`;
    if (rule.operator !== void 0) attrs += ` operator="${DV_OPERATORS[rule.operator]}"`;
    if (rule.allowBlank) attrs += ' allowBlank="1"';
    if (rule.showDropdown === false) attrs += ' showDropDown="1"';
    if (rule.showInputMessage) attrs += ' showInputMessage="1"';
    if (rule.showErrorMessage) attrs += ' showErrorMessage="1"';
    if (rule.errorStyle !== void 0) attrs += ` errorStyle="${rule.errorStyle}"`;
    if (rule.promptTitle !== void 0) attrs += ` promptTitle="${escapeXlsxAttribute(rule.promptTitle)}"`;
    if (rule.prompt !== void 0) attrs += ` prompt="${escapeXlsxAttribute(rule.prompt)}"`;
    if (rule.errorTitle !== void 0) attrs += ` errorTitle="${escapeXlsxAttribute(rule.errorTitle)}"`;
    if (rule.error !== void 0) attrs += ` error="${escapeXlsxAttribute(rule.error)}"`;
    sink.writeString(`<dataValidation${attrs}>`);
    if (rule.values !== void 0) {
      const list = rule.values.map((value) => value.replace(/"/g, '""')).join(",");
      sink.writeString(`<formula1>"${escapeXmlText(list)}"</formula1>`);
    } else if (rule.formula1 !== void 0) {
      sink.writeString(`<formula1>${escapeXmlText(rule.formula1.replace(/^=/, ""))}</formula1>`);
    }
    if (rule.formula2 !== void 0) {
      sink.writeString(`<formula2>${escapeXmlText(rule.formula2.replace(/^=/, ""))}</formula2>`);
    }
    sink.writeString("</dataValidation>");
  }
  sink.writeString("</dataValidations>");
}
function firstCellOfSqref(sqref) {
  const range = sqref.trim().split(/\s+/)[0] ?? "A1";
  return range.split(":")[0] || "A1";
}
function writeConditionalFormatting(sink, groups, startPriority, registerDxf) {
  let priority = startPriority;
  for (const group of groups) {
    sink.writeString(`<conditionalFormatting sqref="${escapeXmlText(group.sqref)}">`);
    const firstCell = firstCellOfSqref(group.sqref);
    for (const rule of group.rules) {
      const current = priority++;
      const dxfIndex = "style" in rule && rule.style ? registerDxf?.(rule.style) ?? -1 : -1;
      const dxfAttr = dxfIndex >= 0 ? ` dxfId="${dxfIndex}"` : "";
      switch (rule.type) {
        case "cellIs": {
          let attrs = ` type="cellIs" operator="${rule.operator}" priority="${current}"${dxfAttr}`;
          if (rule.stopIfTrue) attrs += ' stopIfTrue="1"';
          sink.writeString(`<cfRule${attrs}>`);
          for (const formula of rule.formula) {
            sink.writeString(`<formula>${escapeXmlText(formula.replace(/^=/, ""))}</formula>`);
          }
          sink.writeString("</cfRule>");
          break;
        }
        case "expression": {
          let attrs = ` type="expression" priority="${current}"${dxfAttr}`;
          if (rule.stopIfTrue) attrs += ' stopIfTrue="1"';
          sink.writeString(`<cfRule${attrs}><formula>${escapeXmlText(rule.formula.replace(/^=/, ""))}</formula></cfRule>`);
          break;
        }
        case "colorScale": {
          sink.writeString(`<cfRule type="colorScale" priority="${current}"><colorScale>`);
          for (const point of rule.cfvo) {
            sink.writeString(`<cfvo type="${point.type}"${point.value !== void 0 ? ` val="${escapeXmlText(point.value)}"` : ""}/>`);
          }
          for (const color of rule.colors) {
            sink.writeString(`<color rgb="${colorKey(color).slice(4)}"/>`);
          }
          sink.writeString("</colorScale></cfRule>");
          break;
        }
        case "dataBar": {
          sink.writeString(`<cfRule type="dataBar" priority="${current}"><dataBar${rule.showValue === false ? ' showValue="0"' : ""}>`);
          for (const point of rule.cfvo) {
            sink.writeString(`<cfvo type="${point.type}"${point.value !== void 0 ? ` val="${escapeXmlText(point.value)}"` : ""}/>`);
          }
          sink.writeString(`<color rgb="${colorKey(rule.color).slice(4)}"/></dataBar></cfRule>`);
          break;
        }
        case "iconSet": {
          let attrs = ` type="iconSet" priority="${current}"`;
          if (rule.showValue === false) attrs += ' showValue="0"';
          if (rule.reverse) attrs += ' reverse="1"';
          sink.writeString(`<cfRule${attrs}><iconSet iconSet="${rule.iconSet ?? "3TrafficLights1"}">`);
          for (const point of rule.cfvo ?? []) {
            sink.writeString(`<cfvo type="${point.type}"${point.value !== void 0 ? ` val="${escapeXmlText(point.value)}"` : ""}/>`);
          }
          sink.writeString("</iconSet></cfRule>");
          break;
        }
        case "top10": {
          let attrs = ` type="top10" priority="${current}"${dxfAttr}`;
          if (rule.rank !== void 0) attrs += ` rank="${rule.rank}"`;
          if (rule.percent) attrs += ' percent="1"';
          if (rule.bottom) attrs += ' bottom="1"';
          sink.writeString(`<cfRule${attrs}/>`);
          break;
        }
        case "duplicateValues":
        case "uniqueValues":
          sink.writeString(`<cfRule type="${rule.type}" priority="${current}"${dxfAttr}/>`);
          break;
        case "containsText":
        case "notContainsText":
        case "beginsWith":
        case "endsWith": {
          const operator = rule.type === "containsText" ? "containsText" : rule.type === "notContainsText" ? "notContains" : rule.type === "beginsWith" ? "beginsWith" : "endsWith";
          const escaped = rule.text.replace(/"/g, '""');
          sink.writeString(`<cfRule type="${rule.type}" operator="${operator}" text="${escapeXlsxAttribute(rule.text)}" priority="${current}"${dxfAttr}><formula>${escapeXmlText(`NOT(ISERROR(SEARCH("${escaped}",${firstCell})))`)}</formula></cfRule>`);
          break;
        }
      }
    }
    sink.writeString("</conditionalFormatting>");
  }
}
function writeHyperlinks(sink, links) {
  if (links.length === 0) return;
  sink.writeString("<hyperlinks>");
  for (const { spec, rId } of links) {
    let attrs = ` ref="${escapeXmlText(spec.ref)}"`;
    if (rId) attrs += ` r:id="${rId}"`;
    if (spec.location !== void 0) attrs += ` location="${escapeXlsxAttribute(spec.location)}"`;
    if (spec.tooltip !== void 0) attrs += ` tooltip="${escapeXlsxAttribute(spec.tooltip)}"`;
    if (spec.display !== void 0) attrs += ` display="${escapeXlsxAttribute(spec.display)}"`;
    sink.writeString(`<hyperlink${attrs}/>`);
  }
  sink.writeString("</hyperlinks>");
}
function writeFormulaCell(sink, context, cell, ref, styleAttr, dateStyle, dateTimeStyle) {
  const formula = cell.formula.trim().replace(/^=/, "");
  if (!formula) throw new Error("FormulaCell.formula must not be empty");
  const cached = cell.value;
  let typeAttr = "";
  let cachedXml = "";
  if (cached instanceof ExcelError) {
    typeAttr = ' t="e"';
    cachedXml = `<v>${escapeXmlText(cached.code)}</v>`;
  } else if (cached instanceof Date) {
    const serial = context.toOaDate(cached);
    const effectiveStyle = styleAttr || ` s="${dateHasTime(cached) ? dateTimeStyle : dateStyle}"`;
    sink.writeString(`<c${ref}${effectiveStyle}><f>${escapeXmlText(formula)}</f><v>${serial}</v></c>`);
    return;
  } else if (typeof cached === "number" && Number.isFinite(cached)) {
    cachedXml = `<v>${cached}</v>`;
  } else if (typeof cached === "boolean") {
    typeAttr = ' t="b"';
    cachedXml = `<v>${cached ? 1 : 0}</v>`;
  } else if (cached !== null && cached !== void 0) {
    typeAttr = ' t="str"';
    cachedXml = `<v>${escapeXmlText(String(cached))}</v>`;
  }
  sink.writeString(`<c${ref}${typeAttr}${styleAttr}><f>${escapeXmlText(formula)}</f>${cachedXml}</c>`);
}
function writeXlsxCell(sink, context, raw, col, prevCol, colLetter, rowNum, dateStyle = 1, dateTimeStyle = 2) {
  if (raw === null || raw === void 0) return prevCol;
  const style = getCellStyle(raw);
  const xfIdx = style !== null ? context.registerStyle(style) : -1;
  const styleAttr = xfIdx > 0 ? ` s="${xfIdx}"` : "";
  const ref = cellRefAttr(col, prevCol, colLetter, rowNum);
  const stringStyle = xfIdx > 0 ? xfIdx : void 0;
  const formula = getFormula(raw);
  if (formula !== null) {
    writeFormulaCell(sink, context, raw, ref, styleAttr, dateStyle, dateTimeStyle);
    return col;
  }
  const value = unwrapCell(raw);
  if (value === null || value === void 0) return prevCol;
  if (isRichText(value)) {
    context.writeStringCell(sink, value, ref, stringStyle);
  } else if (typeof value === "number") {
    if (Number.isFinite(value)) {
      sink.writeString(`<c${ref}${styleAttr}><v>${value}</v></c>`);
    } else {
      context.writeStringCell(sink, value.toString(), ref, stringStyle);
    }
  } else if (value instanceof ExcelError) {
    sink.writeString(`<c${ref} t="e"${styleAttr}><v>${escapeXmlText(value.code)}</v></c>`);
  } else if (typeof value === "bigint") {
    context.writeStringCell(sink, value.toString(), ref, stringStyle);
  } else if (typeof value === "boolean") {
    sink.writeString(`<c${ref} t="b"${styleAttr}><v>${value ? 1 : 0}</v></c>`);
  } else if (value instanceof Date) {
    const oaDate = context.toOaDate(value);
    if (Number.isFinite(oaDate)) {
      const finalStyle = xfIdx > 0 ? xfIdx : dateHasTime(value) ? dateTimeStyle : dateStyle;
      sink.writeString(`<c${ref} s="${finalStyle}"><v>${oaDate}</v></c>`);
    } else {
      context.writeStringCell(sink, value.toString(), ref, stringStyle);
    }
  } else {
    context.writeStringCell(sink, value.toString(), ref, stringStyle);
  }
  return col;
}

// src/browser/BrowserXlsxWriter.ts
function headerStyleIndex(style) {
  if (style === "fill") return 4;
  if (style === "bold+fill") return 5;
  return 3;
}
var BrowserXlsxWriter = class {
  constructor(options = {}) {
    this._sheetCount = 0;
    this._sheets = [];
    this._sstValues = [];
    this._sstIndex = /* @__PURE__ */ new Map();
    this._sstCount = 0;
    this._autofilterOn = false;
    this._oaEpoch = Date.UTC(1899, 11, 30);
    this._styles = new XlsxStyleRegistry();
    this._streamBuffer = null;
    this._streamRowNum = 0;
    this._streamColumnCount = 0;
    this._streamDoAutofilter = false;
    this._streamColLetters = [];
    this._streamColWidths = [];
    this._streamPending = null;
    this._streamHeadFlushed = true;
    this._streamHasHeaders = false;
    this._streamDataRows = 0;
    this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    this._zip = new BrowserZip();
    this._useSharedStrings = options.useSharedStrings !== false;
  }
  addSheet(name, hidden = false) {
    const sheetName = this._sanitizeSheetName(name);
    this._sheetCount++;
    const fileName = `sheet${this._sheetCount}.xml`;
    this._sheets.push({
      name: sheetName,
      pathInArchive: `xl/worksheets/${fileName}`,
      nameInArchive: fileName,
      hidden,
      sheetId: this._sheetCount,
      rId: `rId${this._sheetCount}`,
      filterHeaderRange: null
    });
  }
  /** @deprecated Internal; kept for unit-test access. */
  _cl(index) {
    return columnIndexToLetter(index);
  }
  _sanitizeSheetName(name) {
    if (!name || typeof name !== "string") return `Sheet${this._sheetCount + 1}`;
    let value = name.replace(/[\\/*?[\]:]/g, "_");
    if (value.length > 31) value = value.substring(0, 31);
    if (!value.trim()) value = `Sheet${this._sheetCount + 1}`;
    return value;
  }
  // XlsxCellContext ---------------------------------------------------
  writeStringCell(sink, value, ref, styleOverride) {
    assertCellTextLength(typeof value === "string" ? value : value.toString());
    if (this._useSharedStrings) {
      writeXlsxSharedStringCell(
        sink,
        { values: this._sstValues, index: this._sstIndex },
        value,
        ref,
        styleOverride
      );
      this._sstCount++;
      return;
    }
    writeXlsxInlineStringCell(sink, value, ref, styleOverride);
  }
  registerStyle(style) {
    return this._styles.register(style);
  }
  registerFormat(format) {
    return this._styles.registerFormat(format);
  }
  toOaDate(date) {
    return timeToOaDate(date.getTime(), this._oaEpoch);
  }
  // Sheet features/widths ---------------------------------------------
  /**
   * Resolve column widths: explicit `columnWidths` win, otherwise plain
   * autofit from headers and up to 200 sample rows, capped at 200.
   */
  _resolveColumnWidths(columnCount, headers, rows, options, doAutofilter) {
    const widths = initColWidths(columnCount);
    if (options.columnWidths) {
      for (let i = 0; i < columnCount; i++) {
        const width = options.columnWidths[i];
        if (width !== null && width !== void 0) widths[i] = assertColumnWidth(width, i);
      }
      return widths;
    }
    if (options.autofit === false) return widths;
    if (headers) {
      applyHeaderWidths(widths, headers, columnCount, {
        bold: options.headerStyle !== "fill",
        filterButton: doAutofilter
      });
    }
    const samples = options.sampleRows ?? rows;
    const limit = options.sampleRows ? samples?.length ?? 0 : resolveAutofitRows(options.autofitRows);
    if (samples && limit > 0) updateColWidthsFromRows(widths, samples, limit);
    return widths;
  }
  /** Store sheet features and build the worksheet rels/comments parts. */
  _applySheetFeatures(sheet, features) {
    if (features.merges) sheet.merges = features.merges;
    if (features.dataValidations) sheet.dataValidations = features.dataValidations;
    if (features.conditionalFormats) sheet.conditionalFormats = features.conditionalFormats;
    const rels = [];
    let nextRelId = 1;
    const hyperlinks = [];
    for (const spec of features.hyperlinks ?? []) {
      if (spec.target) {
        const id = `rId${nextRelId++}`;
        rels.push({ id, type: XLSX_REL_HYPERLINK, target: spec.target, external: true });
        hyperlinks.push({ spec, rId: id });
      } else {
        hyperlinks.push({ spec });
      }
    }
    if (hyperlinks.length > 0) sheet.hyperlinks = hyperlinks;
    const comments = features.comments ?? [];
    if (comments.length > 0) {
      const commentsPath = `xl/comments/sheet${sheet.sheetId}.xml`;
      const vmlPath = `xl/drawings/commentsDrawing${sheet.sheetId}.vml`;
      const commentsRid = `rId${nextRelId++}`;
      const vmlRid = `rId${nextRelId++}`;
      rels.push({
        id: commentsRid,
        type: XLSX_REL_COMMENTS,
        target: `../comments/sheet${sheet.sheetId}.xml`
      });
      rels.push({
        id: vmlRid,
        type: XLSX_REL_VML,
        target: `../drawings/commentsDrawing${sheet.sheetId}.vml`
      });
      sheet.comments = comments;
      sheet.legacyDrawingRId = vmlRid;
      sheet.extraParts = [
        { name: commentsPath, data: buildXlsxComments(comments) },
        { name: vmlPath, data: buildXlsxCommentsVml(comments) }
      ];
    }
    if (rels.length > 0) sheet.sheetRelsXml = buildXlsxSheetRels(rels);
  }
  _appendExtraParts(sheet) {
    if (sheet.sheetRelsXml) {
      this._zip.addFile(`xl/worksheets/_rels/${sheet.nameInArchive}.rels`, sheet.sheetRelsXml);
    }
    for (const part of sheet.extraParts ?? []) {
      this._zip.addFile(part.name, part.data);
    }
  }
  // Streaming API -----------------------------------------------------
  startSheet(sheetName, columnCount, headers, options = {}) {
    if (this._streamBuffer !== null) throw new Error("Already streaming");
    assertColumnCount(columnCount);
    if (headers !== void 0) {
      assertRowFitsColumns("header", headers.length, columnCount);
      for (const header of headers) assertCellTextLength(header);
    }
    const { hidden = false, doAutofilter = true, headerStyle = "bold" } = options;
    this.addSheet(sheetName, hidden);
    const sheet = this._sheets[this._sheetCount - 1];
    this._applySheetFeatures(sheet, options);
    this._streamBuffer = new BrowserBigBuffer();
    this._streamRowNum = 0;
    this._streamColumnCount = columnCount;
    this._streamDoAutofilter = doAutofilter && headers !== void 0;
    this._streamColWidths = this._resolveColumnWidths(
      columnCount,
      headers,
      void 0,
      options,
      doAutofilter && headers !== void 0
    );
    this._streamColLetters = new Array(columnCount);
    for (let i = 0; i < columnCount; i++) this._streamColLetters[i] = columnIndexToLetter(i);
    this._streamHasHeaders = headers !== void 0;
    this._streamDataRows = 0;
    this._streamAutofitRows = resolveAutofitRows(options.autofitRows);
    this._streamHeadFlushed = false;
    this._streamPending = new BrowserBigBuffer();
    if (headers) {
      const pending = this._streamPending;
      this._streamRowNum++;
      pending.writeString(`<row r="${this._streamRowNum}">`);
      const style = headerStyleIndex(headerStyle);
      for (let c = 0; c < headers.length; c++) {
        this.writeStringCell(pending, headers[c], "", style);
      }
      pending.writeString("</row>");
    }
    if (this._streamAutofitRows === 0) this._flushStreamHead();
  }
  /** Write the delayed sheet head and everything buffered before it. */
  _flushStreamHead() {
    if (this._streamHeadFlushed) return;
    if (this._streamBuffer === null) throw new Error("Not streaming");
    const buffer = this._streamBuffer;
    const pending = this._streamPending;
    writeXlsxSheetHead(buffer, {
      dimensionRef: "A1",
      columnCount: this._streamColumnCount,
      isFirstSheet: this._sheetCount === 1,
      hasHeaders: this._streamHasHeaders,
      colWidths: this._streamColWidths
    });
    if (pending) {
      for (const chunk of pending.getChunks()) buffer.write(chunk);
    }
    this._streamPending = null;
    this._streamHeadFlushed = true;
  }
  _writeStreamRow(buffer, row) {
    this._streamRowNum++;
    assertRowCount(this._streamRowNum);
    buffer.writeString(`<row r="${this._streamRowNum}">`);
    let prevCol = -1;
    for (let c = 0; c < row.length; c++) {
      prevCol = writeXlsxCell(
        buffer,
        this,
        row[c],
        c,
        prevCol,
        this._streamColLetters[c],
        this._streamRowNum
      );
    }
    buffer.writeString("</row>");
  }
  writeRow(row) {
    if (this._streamBuffer === null) throw new Error("Not streaming");
    if (row.length !== this._streamColumnCount) {
      throw new Error(`Row length mismatch. Expected ${this._streamColumnCount} columns, got ${row.length}`);
    }
    if (!this._streamHeadFlushed && this._streamPending) {
      this._streamDataRows++;
      updateColWidthsFromRows(this._streamColWidths, [row]);
      this._writeStreamRow(this._streamPending, row);
      if (this._streamDataRows >= this._streamAutofitRows) this._flushStreamHead();
      return;
    }
    this._writeStreamRow(this._streamBuffer, row);
  }
  endSheet() {
    this._flushStreamHead();
    if (this._streamBuffer === null) throw new Error("Not streaming");
    const buffer = this._streamBuffer;
    const sheet = this._sheets[this._sheetCount - 1];
    const filterRange = writeXlsxSheetTail(buffer, {
      doAutofilter: this._streamDoAutofilter,
      columnCount: this._streamColumnCount,
      lastColumnLetter: this._streamColLetters[this._streamColumnCount - 1] ?? "A",
      rowNum: this._streamRowNum,
      sheetName: sheet.name,
      merges: sheet.merges,
      hyperlinks: sheet.hyperlinks,
      dataValidations: sheet.dataValidations,
      conditionalFormats: sheet.conditionalFormats,
      legacyDrawingRId: sheet.legacyDrawingRId,
      registerDxf: (style) => this._styles.registerDxf(style)
    });
    if (filterRange !== null) {
      this._autofilterOn = true;
      sheet.filterHeaderRange = filterRange;
    }
    this._zip.addFile(sheet.pathInArchive, buffer.toUint8Array());
    this._appendExtraParts(sheet);
    this._streamBuffer = null;
    this._streamPending = null;
    this._streamHeadFlushed = true;
    this._streamHasHeaders = false;
    this._streamDataRows = 0;
    this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    this._streamColWidths = [];
  }
  // Batch API ---------------------------------------------------------
  writeSheet(rows, headers = null, options = {}) {
    const doAutofilter = options.doAutofilter !== false;
    const headerStyle = options.headerStyle ?? "bold";
    const buffer = new BrowserBigBuffer();
    const columnCount = rows.length > 0 ? rows[0].length : headers ? headers.length : 0;
    if (columnCount > 0) assertColumnCount(columnCount);
    const totalRows = rows.length + (headers ? 1 : 0);
    assertRowCount(totalRows);
    if (headers) {
      assertRowFitsColumns("header", headers.length, columnCount);
      for (const header of headers) assertCellTextLength(header);
    }
    for (const row of rows) assertRowFitsColumns("row", row.length, columnCount);
    const sheet = this._sheets[this._sheetCount - 1];
    if (!sheet) throw new Error("Call addSheet() before writeSheet()");
    const colWidths = this._resolveColumnWidths(
      columnCount,
      headers,
      rows,
      options,
      doAutofilter && headers !== null
    );
    const colLetters = new Array(columnCount);
    for (let i = 0; i < columnCount; i++) colLetters[i] = columnIndexToLetter(i);
    const dimensionRef = totalRows > 0 && columnCount > 0 ? `A1:${colLetters[columnCount - 1]}${totalRows}` : "A1";
    writeXlsxSheetHead(buffer, {
      dimensionRef,
      columnCount,
      isFirstSheet: this._sheetCount === 1,
      hasHeaders: headers !== null,
      colWidths
    });
    let rowNum = 0;
    if (headers) {
      rowNum++;
      buffer.writeString(`<row r="${rowNum}">`);
      const style = headerStyleIndex(headerStyle);
      for (let c = 0; c < headers.length; c++) {
        this.writeStringCell(buffer, headers[c], "", style);
      }
      buffer.writeString("</row>");
    }
    for (const row of rows) {
      rowNum++;
      buffer.writeString(`<row r="${rowNum}">`);
      let prevCol = -1;
      for (let c = 0; c < row.length; c++) {
        prevCol = writeXlsxCell(buffer, this, row[c], c, prevCol, colLetters[c], rowNum);
      }
      buffer.writeString("</row>");
    }
    this._applySheetFeatures(sheet, options);
    const filterRange = writeXlsxSheetTail(buffer, {
      doAutofilter: doAutofilter && headers !== null,
      columnCount,
      lastColumnLetter: columnCount > 0 ? colLetters[columnCount - 1] : "A",
      rowNum: totalRows,
      sheetName: sheet.name,
      merges: sheet.merges,
      hyperlinks: sheet.hyperlinks,
      dataValidations: sheet.dataValidations,
      conditionalFormats: sheet.conditionalFormats,
      legacyDrawingRId: sheet.legacyDrawingRId,
      registerDxf: (style) => this._styles.registerDxf(style)
    });
    if (filterRange !== null) {
      this._autofilterOn = true;
      sheet.filterHeaderRange = filterRange;
    }
    this._zip.addFile(sheet.pathInArchive, buffer.toUint8Array());
    this._appendExtraParts(sheet);
  }
  // Package -----------------------------------------------------------
  finalize() {
    this._zip.addFile(
      "xl/sharedStrings.xml",
      buildXlsxSharedStrings(this._sstValues, this._sstCount)
    );
    this._zip.addFile("xl/styles.xml", this._styles.toXml());
    this._zip.addFile("xl/workbook.xml", buildXlsxWorkbook(this._sheets, this._autofilterOn));
    const commentParts = this._sheets.filter((sheet) => (sheet.comments?.length ?? 0) > 0).map((sheet) => ({ path: `xl/comments/sheet${sheet.sheetId}.xml` }));
    this._zip.addFile("[Content_Types].xml", buildXlsxContentTypes(this._sheets, commentParts));
    const rels = buildXlsxRels(this._sheets);
    this._zip.addFile("_rels/.rels", rels.global);
    this._zip.addFile("xl/_rels/workbook.xml.rels", rels.workbook);
    return this._zip.toBlob();
  }
};

// src/biff12Utils.ts
function excelErrorToBiffCode(code) {
  switch (code) {
    case "#NULL!":
      return 0;
    case "#DIV/0!":
      return 7;
    case "#VALUE!":
      return 15;
    case "#REF!":
      return 23;
    case "#NAME?":
      return 29;
    case "#NUM!":
      return 36;
    case "#N/A":
      return 42;
    case "#GETTING_DATA":
      return 43;
    default:
      return 15;
  }
}

// src/writers/core/xlsbRecords.ts
function decodeBase64(base64) {
  const binary = atob(base64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}
var XLSB_STYLES_TEMPLATE = decodeBase64(
  "lgIA5wQEAgAAACwspAATAAAAeQB5AHkAeQBcAC0AbQBtAFwALQBkAGQAXAAgAGgAaAA6AG0AbQAsHqYADAAAAHkAeQB5AHkAXAAtAG0AbQBcAC0AZABkAOgEAOMEBAIAAAArJ9wAAACQAQAAAAIAAAcBAAAAAAD/AgcAAABDAGEAbABpAGIAcgBpACsn3AABALwCAAAAAu4ABwEAAAAAAP8CBwAAAEMAYQBsAGkAYgByAGkAJQYBAAIOAICBCAAmAOQEANsEBAIAAAAtRAAAAAADQAAAAAAA/wNBAAD/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALUQRAAAAA0AAAAAAAP8DQQAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANwEAOUEBAEAAAAuMwAAAAEAAAAAAAAAAAABAAAAAAAAAAAAAQAAAAAAAAAAAAEAAAAAAAAAAAABAAAAAAAAAOYEAPIEBAEAAAAvEP//AAAAAAAAAAAAABAQAADzBADpBAQGAAAALxAAAAAAAAAAAAAAAAAQEAAALxAAAKQAAAAAAAAAAAAQEAEALxAAAKYAAAAAAAAAAAAQEAEALxAAAAAAAQAAAAAAAAAQEAEALxAAAAAAAAABAAAAAAAQEAEALxAAAAAAAQABAAAAAAAQEAEA6gQA6wQEAQAAACUGAQACEQCAgBgQAAAAAAAAAAAAAAAAAAAAACYAMBwAAAAAAQAAAAgAAABOAG8AcgBtAGEAbABuAHkA7AQA+QMEAAAAAPoDAPwDUAAAAAARAAAAVABhAGIAbABlAFMAdAB5AGwAZQBNAGUAZABpAHUAbQAyABEAAABQAGkAdgBvAHQAUwB0AHkAbABlAEwAaQBnAGgAdAAxADYA/QMAIwQCDgAA6wgA9ggqAAAAABEAAABTAGwAaQBjAGUAcgBTAHQAeQBsAGUATABpAGcAaAB0ADEA9wgA7AgAJAAjBAMPAACwEACyEDIAAAAAFQAAAFQAaQBtAGUAUwBsAGkAYwBlAHIAUwB0AHkAbABlAEwAaQBnAGgAdAAxALMQALEQACQAlwIA"
);
var XLSB_BASE_XF_COUNT = 6;
function concat(chunks) {
  let total = 0;
  for (const chunk of chunks) total += chunk.length;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}
function encodeVlq(value) {
  const result = [];
  while (value >= 128) {
    result.push(value & 127 | 128);
    value >>>= 7;
  }
  result.push(value & 127);
  return result;
}
function buildBrtRecord(type, data) {
  const header = [...encodeVlq(type), ...encodeVlq(data.length)];
  return concat([new Uint8Array(header), data]);
}
function writeXlsbRowHeader(sink, rowNum, startCol, endCol) {
  sink.ensureCapacity(27);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(25);
  sink.writeUnsafeInt32LE(rowNum);
  sink.writeUnsafeInt32LE(0);
  sink.writeUnsafeByte(44);
  sink.writeUnsafeByte(1);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(1);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeInt32LE(startCol);
  sink.writeUnsafeInt32LE(endCol);
}
function writeXlsbRkInteger(sink, val, colNum, styleNum = 0) {
  sink.ensureCapacity(14);
  sink.writeUnsafeByte(2);
  sink.writeUnsafeByte(12);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeInt32LE(val << 2 | 2);
}
function writeXlsbDouble(sink, val, colNum, styleNum = 0) {
  sink.ensureCapacity(18);
  sink.writeUnsafeByte(5);
  sink.writeUnsafeByte(16);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeDoubleLE(val);
}
function writeXlsbBool(sink, val, colNum, styleNum = 0) {
  sink.ensureCapacity(13);
  sink.writeUnsafeByte(4);
  sink.writeUnsafeByte(9);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(val ? 1 : 0);
}
function writeXlsbError(sink, code, colNum, styleNum = 0) {
  sink.ensureCapacity(13);
  sink.writeUnsafeByte(3);
  sink.writeUnsafeByte(9);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(excelErrorToBiffCode(code));
}
function writeXlsbIsst(sink, index, colNum, styleNum = 0) {
  sink.ensureCapacity(17);
  sink.writeUnsafeByte(7);
  sink.writeUnsafeByte(12);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeInt32LE(index);
}
function writeXlsbInlineString(sink, value, colNum, styleNum = 0) {
  const cch = value.length;
  const len = 12 + cch * 2;
  sink.ensureCapacity(4 + len);
  sink.writeUnsafeByte(6);
  let length = len;
  while (length >= 128) {
    sink.writeUnsafeByte(length & 127 | 128);
    length >>>= 7;
  }
  sink.writeUnsafeByte(length);
  sink.writeUnsafeInt32LE(colNum);
  sink.writeUnsafeByte(styleNum);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeInt32LE(cch);
  sink.writeUtf16LE(value);
}
function writeXlsbSharedStrings(sink, entries, countAll, countUnique) {
  sink.writeByte(159);
  sink.writeByte(1);
  sink.writeByte(8);
  sink.writeInt32LE(countAll);
  sink.writeInt32LE(countUnique);
  for (const text of entries) {
    const textLength = text.length;
    sink.writeByte(19);
    const recordLength = 5 + 2 * textLength;
    if (recordLength >= 128) {
      sink.writeByte(128 + recordLength % 128);
      const high = recordLength >> 7;
      if (high >= 256) {
        sink.writeByte(128 + high % 128);
      } else {
        sink.writeByte(high);
      }
      sink.writeByte(recordLength >> 14);
      if (recordLength >> 14 > 0) sink.writeByte(0);
    } else {
      sink.writeByte(recordLength & 255);
      sink.writeByte(recordLength >> 8 & 255);
    }
    sink.writeInt32LE(textLength);
    sink.writeUtf16LE(text);
  }
  sink.writeByte(160);
  sink.writeByte(1);
  sink.writeByte(0);
}

// src/writers/core/xlsbStyles.ts
var BRT_FONT = 43;
var BRT_FMT = 44;
var BRT_FILL = 45;
var BRT_BORDER = 46;
var BRT_XF = 47;
var BRT_BEGIN_FMTS = 615;
var BRT_END_FMTS = 616;
var BRT_BEGIN_FONTS = 611;
var BRT_END_FONTS = 612;
var BRT_BEGIN_FILLS = 603;
var BRT_END_FILLS = 604;
var BRT_BEGIN_BORDERS = 613;
var BRT_END_BORDERS = 614;
var BRT_BEGIN_CELL_XFS = 617;
var BRT_END_CELL_XFS = 618;
var BRT_BEGIN_CELL_STYLE_XFS = 626;
var BRT_END_CELL_STYLE_XFS = 627;
var UNDERLINE_BYTES = {
  none: 0,
  single: 1,
  double: 2,
  singleAccounting: 33,
  doubleAccounting: 34
};
var BASELINE_BYTES = {
  baseline: 0,
  superscript: 1,
  subscript: 2
};
var PATTERN_BYTES = {
  none: 0,
  solid: 1,
  mediumGray: 2,
  darkGray: 3,
  lightGray: 4,
  darkHorizontal: 5,
  darkVertical: 6,
  darkDown: 7,
  darkUp: 8,
  darkGrid: 9,
  darkTrellis: 10,
  lightHorizontal: 11,
  lightVertical: 12,
  lightDown: 13,
  lightUp: 14,
  lightGrid: 15,
  lightTrellis: 16,
  gray125: 17,
  gray0625: 18
};
var BORDER_BYTES = {
  none: 0,
  thin: 1,
  medium: 2,
  dashed: 3,
  dotted: 4,
  thick: 5,
  double: 6,
  hair: 7,
  mediumDashed: 8,
  dashDot: 9,
  mediumDashDot: 10,
  dashDotDot: 11,
  mediumDashDotDot: 12,
  slantDashDot: 13
};
var HORIZONTAL_BYTES = {
  general: 0,
  left: 1,
  center: 2,
  right: 3,
  fill: 4,
  justify: 5,
  centerContinuous: 6,
  distributed: 7
};
var VERTICAL_BYTES = {
  top: 0,
  center: 1,
  bottom: 2,
  justify: 3,
  distributed: 4
};
function u16(value) {
  const v = value & 65535;
  return [v & 255, v >>> 8 & 255];
}
function u32(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function utf16leBytes(text) {
  const out = [];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function colorBytes(color) {
  if (color === void 0) return [0, 0, 0, 0, 0, 0, 0, 0];
  if (typeof color === "string") {
    let value = color.trim().replace(/^#/, "").toUpperCase();
    if (value.length === 6) value = `FF${value}`;
    if (!/^[0-9A-F]{8}$/.test(value)) {
      throw new Error(`Invalid color '${color}'. Use RRGGBB, AARRGGBB or '#' + hex.`);
    }
    const alpha = parseInt(value.slice(0, 2), 16);
    const red = parseInt(value.slice(2, 4), 16);
    const green = parseInt(value.slice(4, 6), 16);
    const blue = parseInt(value.slice(6, 8), 16);
    return [5, 0, 0, 0, red, green, blue, alpha];
  }
  if ("indexed" in color) {
    return [3, color.indexed & 255, 0, 0, 0, 0, 0, 255];
  }
  const tint = Math.max(-32768, Math.min(32767, Math.round((color.tint ?? 0) * 32767)));
  return [7, color.theme & 255, ...u16(tint), 0, 0, 0, 255];
}
function buildXlsbFontRecord(font) {
  if (font.name !== void 0 && font.name.length > 31) {
    throw new RangeError("XLSB font names are limited to 31 characters");
  }
  let flags = 0;
  if (font.bold) flags |= 1;
  if (font.italic) flags |= 2;
  if (font.strike) flags |= 8;
  if (font.outline) flags |= 16;
  if (font.shadow) flags |= 32;
  if (font.condense) flags |= 64;
  if (font.extend) flags |= 128;
  const height = Math.max(20, Math.min(8191, Math.round((font.size ?? 11) * 20)));
  const baseline = BASELINE_BYTES[font.vertAlign ?? "baseline"] ?? 0;
  const underline = font.underline === true ? UNDERLINE_BYTES.single : typeof font.underline === "string" ? UNDERLINE_BYTES[font.underline] ?? 0 : 0;
  const payload = [
    ...u16(height),
    ...u16(flags),
    ...u16(font.bold ? 700 : 400),
    ...u16(baseline),
    underline,
    font.family ?? 2,
    font.charset ?? 0,
    0,
    ...colorBytes(font.color)
  ];
  let scheme = font.scheme === "major" ? 1 : font.scheme === "minor" ? 2 : 0;
  if (font.scheme === void 0 && font.name === void 0) {
    scheme = 2;
  }
  payload.push(scheme);
  const name = utf16leBytes(font.name ?? "Calibri");
  payload.push(...u32(name.length / 2), ...name);
  return buildBrtRecord(BRT_FONT, Uint8Array.from(payload));
}
function buildXlsbFillRecord(fill) {
  const payload = [
    ...u32(PATTERN_BYTES[fill.patternType ?? "none"] ?? 0),
    ...colorBytes(fill.fgColor),
    ...colorBytes(fill.bgColor)
  ];
  while (payload.length < 68) payload.push(0);
  return buildBrtRecord(BRT_FILL, Uint8Array.from(payload));
}
function borderSide(side) {
  const style = side?.style !== void 0 ? BORDER_BYTES[side.style] ?? 0 : 0;
  return [style, 0, ...colorBytes(side?.color)];
}
function buildXlsbBorderRecord(border) {
  let diagonal = 0;
  if (border.diagonalDown) diagonal |= 1;
  if (border.diagonalUp) diagonal |= 2;
  const payload = [
    diagonal,
    ...borderSide(border.top),
    ...borderSide(border.bottom),
    ...borderSide(border.left),
    ...borderSide(border.right),
    ...borderSide(border.diagonal)
  ];
  return buildBrtRecord(BRT_BORDER, Uint8Array.from(payload));
}
function buildXlsbXfRecord(options) {
  const { ifmt = 0, fontId = 0, fillId = 0, borderId = 0, style } = options;
  const alignment = style?.alignment;
  const protection = style?.protection;
  let rotation = 0;
  let indent = 0;
  let alignmentByte = 16;
  let flagsByte = 16;
  if (alignment) {
    rotation = alignment.textRotation ?? 0;
    indent = alignment.indent ?? 0;
    alignmentByte = (HORIZONTAL_BYTES[alignment.horizontal ?? "general"] ?? 0) | (VERTICAL_BYTES[alignment.vertical ?? "bottom"] ?? 2) << 3 | (alignment.wrapText ? 64 : 0) | (alignment.justifyLastLine ? 128 : 0);
    flagsByte = 16 | (alignment.shrinkToFit ? 1 : 0) | (alignment.readingOrder ?? 0) << 2;
  }
  if (protection) {
    flagsByte = flagsByte & 207 | (protection.locked !== false ? 16 : 0) | (protection.hidden ? 32 : 0);
  }
  let flags = 0;
  if (ifmt !== 0) flags |= 1;
  if (fontId !== 0) flags |= 2;
  if (alignment) flags |= 4;
  if (borderId !== 0) flags |= 8;
  if (fillId !== 0) flags |= 16;
  if (protection) flags |= 32;
  const payload = [
    0,
    0,
    // ixfeParent: built-in cell style XF 0
    ...u16(ifmt),
    ...u16(fontId),
    ...u16(fillId),
    ...u16(borderId),
    rotation & 255,
    indent & 255,
    alignmentByte & 255,
    flagsByte & 255,
    ...u16(flags)
  ];
  return buildBrtRecord(BRT_XF, Uint8Array.from(payload));
}
function numFmtRecord(ifmt, format) {
  if (format.length > 255) {
    throw new RangeError("XLSB number formats are limited to 255 characters");
  }
  const encoded = utf16leBytes(format);
  const payload = [...u16(ifmt), ...u32(encoded.length / 2), ...encoded];
  return buildBrtRecord(BRT_FMT, Uint8Array.from(payload));
}
function scanRecords(data) {
  const spans = [];
  let pos = 0;
  while (pos < data.length) {
    const start = pos;
    let id = 0;
    let shift = 0;
    let byte;
    do {
      if (pos >= data.length) return spans;
      byte = data[pos++];
      id |= (byte & 127) << shift;
      shift += 7;
    } while (byte & 128);
    let len = 0;
    shift = 0;
    do {
      if (pos >= data.length) return spans;
      byte = data[pos++];
      len |= (byte & 127) << shift;
      shift += 7;
    } while (byte & 128);
    if (pos + len > data.length) return spans;
    spans.push({ id, start, dataStart: pos, end: pos + len });
    pos += len;
  }
  return spans;
}
function countRecord(id, count) {
  return buildBrtRecord(id, Uint8Array.from(u32(count)));
}
var XlsbStyleRegistry = class {
  constructor() {
    this.fontIds = /* @__PURE__ */ new Map();
    this.fillIds = /* @__PURE__ */ new Map();
    this.borderIds = /* @__PURE__ */ new Map();
    this.xfIds = /* @__PURE__ */ new Map();
    this.formats = /* @__PURE__ */ new Map();
    this.fonts = [];
    this.fills = [];
    this.borders = [];
    this.xfs = [];
    this.formatRecords = [];
    this.nextNumFmtId = 167;
    this.nextFontId = 2;
    this.nextFillId = 2;
    this.nextBorderId = 1;
  }
  /** Cell XF index for a style; 0 for null/empty styles. */
  register(style) {
    const normalized = normalizeStyle(style);
    if (normalized === null) return 0;
    const key = JSON.stringify(normalized);
    const existing = this.xfIds.get(key);
    if (existing !== void 0) return existing;
    const base = baseXfIndex(normalized);
    if (base !== null) {
      this.xfIds.set(key, base);
      return base;
    }
    const fontId = this.componentId(
      this.fontIds,
      normalized.font ? JSON.stringify(normalized.font) : null,
      normalized.font ? buildXlsbFontRecord(normalized.font) : null,
      this.fonts,
      normalized.font?.bold && Object.keys(normalized.font).length === 1 ? 1 : null,
      () => this.nextFontId++
    );
    const fillId = this.componentId(
      this.fillIds,
      normalized.fill ? JSON.stringify(normalized.fill) : null,
      normalized.fill ? buildXlsbFillRecord(normalized.fill) : null,
      this.fills,
      null,
      () => this.nextFillId++
    );
    const borderId = this.componentId(
      this.borderIds,
      normalized.border ? JSON.stringify(normalized.border) : null,
      normalized.border ? buildXlsbBorderRecord(normalized.border) : null,
      this.borders,
      null,
      () => this.nextBorderId++
    );
    let ifmt = 0;
    if (normalized.numFmt !== void 0) {
      ifmt = this.numFmtIdFor(normalized.numFmt);
    }
    const index = XLSB_BASE_XF_COUNT + this.xfs.length;
    this.xfs.push(buildXlsbXfRecord({
      ifmt,
      fontId,
      fillId,
      borderId,
      style: normalized
    }));
    this.xfIds.set(key, index);
    return index;
  }
  /** XF index for a plain number format (legacy FormattedCell path). */
  registerFormat(format) {
    return this.register({ numFmt: format });
  }
  numFmtIdFor(format) {
    let id = this.formats.get(format);
    if (id === void 0) {
      if (this.nextNumFmtId > 382) {
        throw new RangeError("XLSB supports at most 216 custom number formats");
      }
      id = this.nextNumFmtId++;
      this.formats.set(format, id);
      this.formatRecords.push({ ifmt: id, format });
    }
    return id;
  }
  componentId(registry, key, record, target, baseId, allocate) {
    if (key === null) return 0;
    const existing = registry.get(key);
    if (existing !== void 0) return existing;
    if (baseId !== null) {
      registry.set(key, baseId);
      return baseId;
    }
    const id = allocate();
    registry.set(key, id);
    if (record) target.push(record);
    return id;
  }
  /** Rewrite `styles.bin`, appending registered components to the template. */
  toStylesBin(template) {
    if (this.xfs.length === 0 && this.fonts.length === 0 && this.fills.length === 0 && this.borders.length === 0 && this.formatRecords.length === 0) {
      return template;
    }
    const records = scanRecords(template);
    const find = (id) => {
      const record = records.find((r) => r.id === id);
      if (!record) throw new Error(`XLSB styles template is missing record 0x${id.toString(16)}`);
      return record;
    };
    const section = (beginId, endId) => {
      const begin = find(beginId);
      const end = find(endId);
      const inner = records.filter((r) => r.start >= begin.end && r.end <= end.start);
      return { begin, end, inner };
    };
    const parts = [];
    const fmts = section(BRT_BEGIN_FMTS, BRT_END_FMTS);
    parts.push(template.subarray(0, fmts.begin.start));
    parts.push(countRecord(BRT_BEGIN_FMTS, fmts.inner.length + this.formatRecords.length));
    for (const record of fmts.inner) parts.push(template.subarray(record.start, record.end));
    for (const fmt of this.formatRecords) parts.push(numFmtRecord(fmt.ifmt, fmt.format));
    parts.push(template.subarray(fmts.end.start, fmts.end.end));
    appendComponentSection(
      parts,
      template,
      section(BRT_BEGIN_FONTS, BRT_END_FONTS),
      BRT_FONT,
      this.fonts
    );
    appendComponentSection(
      parts,
      template,
      section(BRT_BEGIN_FILLS, BRT_END_FILLS),
      BRT_FILL,
      this.fills
    );
    appendComponentSection(
      parts,
      template,
      section(BRT_BEGIN_BORDERS, BRT_END_BORDERS),
      BRT_BORDER,
      this.borders
    );
    const styleXfs = find(BRT_BEGIN_CELL_STYLE_XFS);
    const styleXfsEnd = find(BRT_END_CELL_STYLE_XFS);
    parts.push(template.subarray(styleXfs.start, styleXfsEnd.end));
    const cellXfs = section(BRT_BEGIN_CELL_XFS, BRT_END_CELL_XFS);
    parts.push(countRecord(BRT_BEGIN_CELL_XFS, cellXfs.inner.length + this.xfs.length));
    for (const record of cellXfs.inner) parts.push(template.subarray(record.start, record.end));
    for (const xf of this.xfs) parts.push(xf);
    parts.push(template.subarray(cellXfs.end.start, cellXfs.end.end));
    parts.push(template.subarray(cellXfs.end.end));
    return concat(parts);
  }
};
function appendComponentSection(parts, template, section, componentId, custom) {
  parts.push(countRecord(section.begin.id, section.inner.length + custom.length));
  let splitAt = section.begin.end;
  for (const record of section.inner) {
    if (record.id === componentId) splitAt = record.end;
  }
  parts.push(template.subarray(section.begin.end, splitAt));
  for (const record of custom) parts.push(record);
  parts.push(template.subarray(splitAt, section.end.start));
  parts.push(template.subarray(section.end.start, section.end.end));
}
function baseXfIndex(style) {
  const keys = Object.keys(style);
  if (keys.length === 1 && keys[0] === "font" && style.font?.bold === true && Object.keys(style.font).length === 1) {
    return 3;
  }
  if (keys.length === 1 && keys[0] === "fill" && style.fill?.patternType === "gray125" && Object.keys(style.fill).length === 1) {
    return 4;
  }
  if (keys.length === 2 && style.font?.bold === true && Object.keys(style.font).length === 1 && style.fill?.patternType === "gray125" && Object.keys(style.fill).length === 1) {
    return 5;
  }
  return null;
}

// src/writers/core/xlsbFormula.ts
var BRT_FMLA_STRING = 8;
var BRT_FMLA_NUM = 9;
var BRT_FMLA_BOOL = 10;
var BRT_FMLA_ERROR = 11;
var FUNCTIONS = {
  COUNT: { id: 0, min: 0, max: 255, fixed: null, classes: "R" },
  IF: { id: 1, min: 2, max: 3, fixed: null, classes: "VRR" },
  ISNA: { id: 2, min: 1, max: 1, fixed: 1, classes: "V" },
  ISERROR: { id: 3, min: 1, max: 1, fixed: 1, classes: "V" },
  SUM: { id: 4, min: 0, max: 255, fixed: null, classes: "R" },
  AVERAGE: { id: 5, min: 1, max: 255, fixed: null, classes: "R" },
  MIN: { id: 6, min: 1, max: 255, fixed: null, classes: "R" },
  MAX: { id: 7, min: 1, max: 255, fixed: null, classes: "R" },
  ROW: { id: 8, min: 0, max: 1, fixed: null, classes: "R" },
  COLUMN: { id: 9, min: 0, max: 1, fixed: null, classes: "R" },
  NA: { id: 10, min: 0, max: 0, fixed: 0, classes: "" },
  STDEV: { id: 12, min: 1, max: 255, fixed: null, classes: "R" },
  SIN: { id: 15, min: 1, max: 1, fixed: 1, classes: "V" },
  COS: { id: 16, min: 1, max: 1, fixed: 1, classes: "V" },
  TAN: { id: 17, min: 1, max: 1, fixed: 1, classes: "V" },
  ATAN: { id: 18, min: 1, max: 1, fixed: 1, classes: "V" },
  PI: { id: 19, min: 0, max: 0, fixed: 0, classes: "" },
  SQRT: { id: 20, min: 1, max: 1, fixed: 1, classes: "V" },
  EXP: { id: 21, min: 1, max: 1, fixed: 1, classes: "V" },
  LN: { id: 22, min: 1, max: 1, fixed: 1, classes: "V" },
  LOG10: { id: 23, min: 1, max: 1, fixed: 1, classes: "V" },
  ABS: { id: 24, min: 1, max: 1, fixed: 1, classes: "V" },
  INT: { id: 25, min: 1, max: 1, fixed: 1, classes: "V" },
  SIGN: { id: 26, min: 1, max: 1, fixed: 1, classes: "V" },
  ROUND: { id: 27, min: 2, max: 2, fixed: 2, classes: "VV" },
  INDEX: { id: 29, min: 2, max: 4, fixed: null, classes: "RVVV" },
  MID: { id: 31, min: 3, max: 3, fixed: 3, classes: "VVV" },
  LEN: { id: 32, min: 1, max: 1, fixed: 1, classes: "V" },
  VALUE: { id: 33, min: 1, max: 1, fixed: 1, classes: "V" },
  AND: { id: 36, min: 1, max: 255, fixed: null, classes: "R" },
  OR: { id: 37, min: 1, max: 255, fixed: null, classes: "R" },
  NOT: { id: 38, min: 1, max: 1, fixed: 1, classes: "V" },
  MOD: { id: 39, min: 2, max: 2, fixed: 2, classes: "VV" },
  VAR: { id: 46, min: 1, max: 255, fixed: null, classes: "R" },
  TEXT: { id: 48, min: 2, max: 2, fixed: 2, classes: "VV" },
  MATCH: { id: 64, min: 2, max: 3, fixed: null, classes: "VRR" },
  DATE: { id: 65, min: 3, max: 3, fixed: 3, classes: "VVV" },
  DAY: { id: 67, min: 1, max: 1, fixed: 1, classes: "V" },
  MONTH: { id: 68, min: 1, max: 1, fixed: 1, classes: "V" },
  YEAR: { id: 69, min: 1, max: 1, fixed: 1, classes: "V" },
  HOUR: { id: 71, min: 1, max: 1, fixed: 1, classes: "V" },
  MINUTE: { id: 72, min: 1, max: 1, fixed: 1, classes: "V" },
  SECOND: { id: 73, min: 1, max: 1, fixed: 1, classes: "V" },
  NOW: { id: 74, min: 0, max: 0, fixed: 0, classes: "" },
  ROWS: { id: 76, min: 1, max: 1, fixed: 1, classes: "R" },
  COLUMNS: { id: 77, min: 1, max: 1, fixed: 1, classes: "R" },
  SEARCH: { id: 82, min: 2, max: 3, fixed: null, classes: "VVV" },
  HLOOKUP: { id: 101, min: 3, max: 4, fixed: null, classes: "VRRV" },
  VLOOKUP: { id: 102, min: 3, max: 4, fixed: null, classes: "VRRV" },
  LOG: { id: 109, min: 1, max: 2, fixed: null, classes: "VV" },
  LOWER: { id: 112, min: 1, max: 1, fixed: 1, classes: "V" },
  UPPER: { id: 113, min: 1, max: 1, fixed: 1, classes: "V" },
  LEFT: { id: 115, min: 1, max: 2, fixed: null, classes: "VV" },
  RIGHT: { id: 116, min: 1, max: 2, fixed: null, classes: "VV" },
  EXACT: { id: 117, min: 2, max: 2, fixed: 2, classes: "VV" },
  TRIM: { id: 118, min: 1, max: 1, fixed: 1, classes: "V" },
  REPLACE: { id: 119, min: 4, max: 4, fixed: 4, classes: "VVVV" },
  SUBSTITUTE: { id: 120, min: 3, max: 4, fixed: null, classes: "VVVV" },
  FIND: { id: 124, min: 2, max: 3, fixed: null, classes: "VVV" },
  ISTEXT: { id: 127, min: 1, max: 1, fixed: 1, classes: "V" },
  ISNUMBER: { id: 128, min: 1, max: 1, fixed: 1, classes: "V" },
  ISBLANK: { id: 129, min: 1, max: 1, fixed: 1, classes: "V" },
  T: { id: 130, min: 1, max: 1, fixed: 1, classes: "R" },
  N: { id: 131, min: 1, max: 1, fixed: 1, classes: "R" },
  TRUNC: { id: 197, min: 1, max: 2, fixed: null, classes: "VV" },
  ROUNDUP: { id: 212, min: 2, max: 2, fixed: 2, classes: "VV" },
  ROUNDDOWN: { id: 213, min: 2, max: 2, fixed: 2, classes: "VV" },
  RANK: { id: 216, min: 2, max: 3, fixed: null, classes: "VRV" },
  TODAY: { id: 221, min: 0, max: 0, fixed: 0, classes: "" },
  SUMPRODUCT: { id: 228, min: 1, max: 255, fixed: null, classes: "A" },
  FLOOR: { id: 285, min: 2, max: 2, fixed: 2, classes: "VV" },
  CEILING: { id: 288, min: 2, max: 2, fixed: 2, classes: "VV" },
  CONCATENATE: { id: 336, min: 0, max: 255, fixed: null, classes: "V" },
  POWER: { id: 337, min: 2, max: 2, fixed: 2, classes: "VV" },
  SUMIF: { id: 345, min: 2, max: 3, fixed: null, classes: "RVR" },
  COUNTIF: { id: 346, min: 2, max: 2, fixed: 2, classes: "RV" },
  COUNTBLANK: { id: 347, min: 1, max: 1, fixed: 1, classes: "R" },
  AVERAGEA: { id: 361, min: 1, max: 255, fixed: null, classes: "R" },
  MAXA: { id: 362, min: 1, max: 255, fixed: null, classes: "R" },
  MINA: { id: 363, min: 1, max: 255, fixed: null, classes: "R" },
  IFERROR: { id: 480, min: 2, max: 2, fixed: 2, classes: "VV" },
  COUNTIFS: { id: 481, min: 2, max: 255, fixed: null, classes: "RV" },
  SUMIFS: { id: 482, min: 3, max: 255, fixed: null, classes: "RRV" },
  AVERAGEIF: { id: 483, min: 2, max: 3, fixed: null, classes: "RVR" }
};
var VOLATILE = /* @__PURE__ */ new Set(["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"]);
var OPERATOR_IDS = {
  "+": 3,
  "-": 4,
  "*": 5,
  "/": 6,
  "^": 7,
  "&": 8,
  "<": 9,
  "<=": 10,
  "=": 11,
  ">=": 12,
  ">": 13,
  "<>": 14,
  " ": 15,
  ",": 16,
  ":": 17
};
var PRECEDENCE = {
  ",": 1,
  "=": 2,
  "<": 2,
  ">": 2,
  "<=": 2,
  ">=": 2,
  "<>": 2,
  "&": 3,
  "+": 4,
  "-": 4,
  "*": 5,
  "/": 5,
  "^": 6,
  " ": 9,
  ":": 10
};
var ERROR_CODES = {
  "#NULL!": 0,
  "#DIV/0!": 7,
  "#VALUE!": 15,
  "#REF!": 23,
  "#NAME?": 29,
  "#NUM!": 36,
  "#N/A": 42,
  "#GETTING_DATA": 43
};
var ADDRESS = String.raw`(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)`;
var TOKEN_RE = new RegExp(
  String.raw`\s+|"(?:[^"]|"")*"|'(?:[^']|'')*'` + String.raw`|#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)` + `|${ADDRESS}(?![\\w.])` + String.raw`|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[^\W\d][\w.]*|<>|<=|>=|[+\-*/^&%=<>(),:]`,
  "giu"
);
function tokenize(text) {
  const tokens = [];
  let position = 0;
  while (position < text.length) {
    TOKEN_RE.lastIndex = position;
    const match = TOKEN_RE.exec(text);
    if (!match || match.index !== position) {
      throw new Error(`Unsupported XLSB formula syntax near '${text.slice(position)}'`);
    }
    position = TOKEN_RE.lastIndex;
    const token = match[0];
    if (/^\s+$/.test(token)) continue;
    tokens.push(token);
  }
  return tokens;
}
var Parser = class {
  constructor(tokens) {
    this.tokens = tokens;
    this.index = 0;
  }
  peek() {
    return this.tokens[this.index];
  }
  pop() {
    const token = this.tokens[this.index++];
    if (token === void 0) throw new Error("Unexpected end of XLSB formula");
    return token;
  }
  parse() {
    const node = this.expression(0);
    if (this.peek() !== void 0) {
      throw new Error(`Unsupported XLSB formula token '${this.peek()}'`);
    }
    return node;
  }
  expression(minPrecedence) {
    let left = this.primary();
    while (this.peek() === "%") {
      this.pop();
      left = { kind: "percent", value: left };
    }
    for (; ; ) {
      const operator = this.peek();
      if (operator === void 0) break;
      const precedence = PRECEDENCE[operator];
      if (precedence === void 0 || precedence < minPrecedence) break;
      this.pop();
      const right = this.expression(precedence + 1);
      left = { kind: "binary", op: operator, left, right };
    }
    return left;
  }
  primary() {
    const token = this.pop();
    if (token === "+" || token === "-") {
      return { kind: "unary", op: token, value: this.expression(7) };
    }
    if (token === "(") {
      const value = this.expression(0);
      if (this.pop() !== ")") throw new Error("Missing ) in XLSB formula");
      return { kind: "paren", value };
    }
    if (token.startsWith('"')) {
      return { kind: "string", value: token.slice(1, -1).replace(/""/g, '"') };
    }
    if (token.startsWith("#")) {
      const code = ERROR_CODES[token.toUpperCase()];
      if (code === void 0) throw new Error(`Unsupported XLSB formula error literal '${token}'`);
      return { kind: "error", value: code };
    }
    if (/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(token)) {
      return { kind: "number", value: Number(token) };
    }
    if (/^[A-Za-z_\\]/.test(token)) {
      const upper = token.toUpperCase();
      if (this.peek() === "(") {
        this.pop();
        const args = [];
        if (this.peek() !== ")") {
          for (; ; ) {
            args.push(this.peek() === "," || this.peek() === ")" ? { kind: "missing" } : this.expression(2));
            if (this.peek() !== ",") break;
            this.pop();
          }
        }
        if (this.pop() !== ")") throw new Error("Missing ) in XLSB function call");
        return { kind: "function", name: upper, args };
      }
      if (upper === "TRUE" || upper === "FALSE") {
        return { kind: "bool", value: upper === "TRUE" };
      }
      if (ADDRESS_RE.test(token)) {
        return { kind: "reference", value: token };
      }
      throw new Error(`Unsupported XLSB formula name '${token}' (defined names are not supported)`);
    }
    if (ADDRESS_RE.test(token)) {
      return { kind: "reference", value: token };
    }
    throw new Error(`Unsupported XLSB formula token '${token}'`);
  }
};
var ADDRESS_RE = new RegExp(`^(?:${ADDRESS})$`, "i");
function columnIndex(letters) {
  let value = 0;
  for (const ch of letters.toUpperCase()) {
    value = value * 26 + (ch.charCodeAt(0) - 64);
  }
  return value - 1;
}
function parseReferencePart(text, kind) {
  if (kind === "cell") {
    const match2 = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(text);
    if (!match2) throw new Error(`Invalid XLSB cell reference '${text}'`);
    const col = columnIndex(match2[2]);
    const row = parseInt(match2[4], 10) - 1;
    if (col > 16383 || row > 1048575) {
      throw new RangeError(`Cell reference '${text}' is outside the Excel grid`);
    }
    return { row, col, colRelative: match2[1] === "", rowRelative: match2[3] === "" };
  }
  if (kind === "col") {
    const match2 = /^(\$?)([A-Za-z]{1,3})$/.exec(text);
    if (!match2) throw new Error(`Invalid XLSB column reference '${text}'`);
    return { row: 0, col: columnIndex(match2[2]), rowRelative: false, colRelative: match2[1] === "" };
  }
  const match = /^(\$?)([1-9]\d*)$/.exec(text);
  if (!match) throw new Error(`Invalid XLSB row reference '${text}'`);
  return { row: parseInt(match[2], 10) - 1, col: 0, rowRelative: match[1] === "", colRelative: false };
}
function refKind(part) {
  if (/^[A-Za-z]/.test(part.replace(/^\$/, ""))) {
    return /^\d|^\$\d/.test(part.replace(/^\$?[A-Za-z]{1,3}/, "")) ? "cell" : "col";
  }
  return "row";
}
function columnField(part) {
  return part.col & 16383 | (part.colRelative ? 16384 : 0) | (part.rowRelative ? 32768 : 0);
}
function emitReference(text, operand) {
  const parts = text.split(":");
  const base = operand === "R" ? 32 : operand === "A" ? 96 : 64;
  if (parts.length === 1) {
    const part = parseReferencePart(parts[0], "cell");
    return [base | 4, ...u322(part.row), ...u162(columnField(part))];
  }
  if (parts.length !== 2) throw new Error(`Invalid XLSB range '${text}'`);
  const kind = refKind(parts[0]);
  if (kind !== refKind(parts[1])) {
    throw new Error(`Invalid XLSB range '${text}'`);
  }
  let first = parseReferencePart(parts[0], kind);
  let last = parseReferencePart(parts[1], kind);
  if (kind === "col") {
    first = { ...first, row: 0, rowRelative: false };
    last = { ...last, row: 1048575, rowRelative: false };
  } else if (kind === "row") {
    first = { ...first, col: 0, colRelative: false };
    last = { ...last, col: 16383, colRelative: false };
  }
  return [
    base | 5,
    ...u322(first.row),
    ...u322(last.row),
    ...u162(columnField(first)),
    ...u162(columnField(last))
  ];
}
function u162(value) {
  const v = value & 65535;
  return [v & 255, v >>> 8 & 255];
}
function u322(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function doubleBytes(value) {
  const buffer = new Uint8Array(8);
  new DataView(buffer.buffer).setFloat64(0, value, true);
  return [...buffer];
}
function stringBytes(text) {
  if (text.length > 255) {
    throw new RangeError("XLSB formula string literals are limited to 255 UTF-16 units");
  }
  const out = [...u162(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function functionArgumentClass(info, index) {
  if (info.id === 481) return index % 2 === 0 ? "R" : "V";
  if (info.id === 482 || info.id === 484) {
    return index === 0 || index % 2 === 0 ? "R" : "V";
  }
  const chars = info.classes;
  if (chars.length === 0) return "V";
  return chars[Math.min(index, chars.length - 1)] === "R" ? "R" : "V";
}
function emitNode(node, operand = "V") {
  switch (node.kind) {
    case "number": {
      if (Number.isInteger(node.value) && node.value >= 0 && node.value <= 65535) {
        return [30, ...u162(node.value)];
      }
      return [31, ...doubleBytes(node.value)];
    }
    case "string":
      return [23, ...stringBytes(node.value)];
    case "bool":
      return [29, node.value ? 1 : 0];
    case "error":
      return [28, node.value];
    case "missing":
      return [22];
    case "reference":
      return emitReference(node.value, operand);
    case "paren":
      return [...emitNode(node.value, operand), 21];
    case "percent":
      return [...emitNode(node.value), 20];
    case "unary":
      return [...emitNode(node.value), node.op === "+" ? 18 : 19];
    case "binary": {
      const argType = node.op === "," || node.op === " " || node.op === ":" ? "R" : "V";
      return [
        ...emitNode(node.left, argType),
        ...emitNode(node.right, argType),
        OPERATOR_IDS[node.op]
      ];
    }
    case "function": {
      const info = FUNCTIONS[node.name];
      if (!info) {
        throw new Error(`Unsupported XLSB function '${node.name}'`);
      }
      if (node.args.length < info.min || node.args.length > info.max) {
        throw new RangeError(
          `${node.name} expects ${info.min}..${info.max} arguments, got ${node.args.length}`
        );
      }
      const encoded = node.args.map((arg, index) => emitNode(arg, functionArgumentClass(info, index)));
      const resultClass = operand === "R" ? 32 : operand === "A" ? 96 : 64;
      const ending = info.fixed !== null ? [resultClass | 1, ...u162(info.id)] : [resultClass | 2, node.args.length, ...u162(info.id)];
      if (info.id === 1) {
        const yes = encoded[1] ?? [];
        const no = encoded[2] ?? [];
        const firstGoto = [25, 8, ...u162(no.length + (no.length > 0 ? 4 : 0) + ending.length - 1)];
        const lastGoto = no.length > 0 ? [25, 8, ...u162(ending.length - 1)] : [];
        return [
          ...encoded[0] ?? [],
          25,
          2,
          ...u162(yes.length + 4),
          ...yes,
          ...firstGoto,
          ...no,
          ...lastGoto,
          ...ending
        ];
      }
      if (info.id === 480) {
        const fallback = encoded[1] ?? [];
        return [
          ...encoded[0] ?? [],
          25,
          128,
          ...u162(fallback.length),
          ...fallback,
          25,
          8,
          ...u162(ending.length - 1),
          ...ending
        ];
      }
      return [...encoded.flat(), ...ending];
    }
  }
}
function containsVolatile(node) {
  switch (node.kind) {
    case "function":
      return VOLATILE.has(node.name) || node.args.some(containsVolatile);
    case "paren":
    case "percent":
    case "unary":
      return containsVolatile(node.value);
    case "binary":
      return containsVolatile(node.left) || containsVolatile(node.right);
    default:
      return false;
  }
}
function compileXlsbFormula(formula) {
  const text = formula.trim().replace(/^=/, "");
  if (!text) throw new Error("FormulaCell.formula must not be empty");
  if (text.length > 8192) throw new RangeError("XLSB formula text is limited to 8192 characters");
  const tree = new Parser(tokenize(text)).parse();
  let tokens = emitNode(tree);
  if (containsVolatile(tree)) {
    tokens = [25, 1, 0, 0, ...tokens];
  }
  if (tokens.length === 0 || tokens.length > 16384) {
    throw new RangeError("XLSB formula token length exceeds the BIFF12 limit");
  }
  return Uint8Array.from(tokens);
}
function wideString(text) {
  const out = [...u322(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function buildXlsbFormulaRecord(column, styleIndex, formula, cached) {
  let recordId = BRT_FMLA_NUM;
  let cache;
  if (cached instanceof ExcelError) {
    recordId = BRT_FMLA_ERROR;
    cache = [excelErrorToBiffCode(cached.code)];
  } else if (typeof cached === "string") {
    recordId = BRT_FMLA_STRING;
    cache = wideString(cached);
  } else if (typeof cached === "boolean") {
    recordId = BRT_FMLA_BOOL;
    cache = [cached ? 1 : 0];
  } else {
    const value = typeof cached === "number" ? cached : 0;
    if (!Number.isFinite(value)) {
      throw new RangeError("XLSB formula cached values must be finite");
    }
    cache = doubleBytes(value);
  }
  const tokens = compileXlsbFormula(formula);
  const payload = [
    ...u322(column),
    ...u322(styleIndex),
    ...cache,
    1,
    0,
    ...u322(tokens.length),
    ...tokens,
    ...u322(0)
  ];
  return buildBrtRecord(recordId, Uint8Array.from(payload));
}

// src/writers/core/xlsbSheetFeatures.ts
var BRT_BEGIN_MERGE_CELLS = 177;
var BRT_MERGE_CELL = 176;
var BRT_END_MERGE_CELLS = 178;
var BRT_BEGIN_DVALS = 573;
var BRT_DVAL = 64;
var BRT_END_DVALS = 574;
var BRT_HLINK = 494;
var DV_TYPES2 = {
  whole: 1,
  decimal: 2,
  list: 3,
  date: 4,
  time: 5,
  textLength: 6,
  custom: 7
};
var DV_ERROR_STYLES = {
  stop: 0,
  warning: 1,
  information: 2
};
var DV_OPERATORS2 = {
  between: 0,
  notBetween: 1,
  equal: 2,
  notEqual: 3,
  lessThan: 4,
  lessThanOrEqual: 5,
  greaterThan: 6,
  greaterThanOrEqual: 7
};
function u163(value) {
  const v = value & 65535;
  return [v & 255, v >>> 8 & 255];
}
function u323(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function wideString2(text) {
  const out = [...u323(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function nullableString(text) {
  return text === void 0 || text === "" ? [255, 255, 255, 255] : wideString2(text);
}
function columnIndex2(letters) {
  let value = 0;
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value - 1;
}
function parseCell(text) {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(text.trim());
  if (!match) throw new Error(`Invalid cell reference '${text}'`);
  const col = columnIndex2(match[1]);
  const row = parseInt(match[2], 10) - 1;
  if (col < 0 || col > 16383 || row < 0 || row > 1048575) {
    throw new RangeError(`Cell reference '${text}' is outside the Excel grid`);
  }
  return { col, row };
}
function parseXlsbRange(text) {
  const parts = text.split(":");
  if (parts.length > 2) throw new Error(`Invalid range '${text}'`);
  const first = parseCell(parts[0]);
  const last = parts.length === 2 ? parseCell(parts[1]) : first;
  return {
    c1: Math.min(first.col, last.col),
    r1: Math.min(first.row, last.row),
    c2: Math.max(first.col, last.col),
    r2: Math.max(first.row, last.row)
  };
}
function buildXlsbMergeRecords(merges) {
  if (merges.length === 0) return new Uint8Array(0);
  const parts = [
    buildBrtRecord(BRT_BEGIN_MERGE_CELLS, Uint8Array.from(u323(merges.length)))
  ];
  for (const ref of merges) {
    const range = parseXlsbRange(ref);
    parts.push(buildBrtRecord(BRT_MERGE_CELL, Uint8Array.from([
      ...u323(range.r1),
      ...u323(range.r2),
      ...u323(range.c1),
      ...u323(range.c2)
    ])));
  }
  parts.push(buildBrtRecord(BRT_END_MERGE_CELLS, new Uint8Array(0)));
  return concat(parts);
}
function buildXlsbHyperlinkRecord(row, col, link, rId) {
  const internal = link.target === void 0;
  const relationship = internal ? "" : rId ?? "";
  const location = internal ? link.location ?? "" : "";
  const display = link.display ?? (internal ? link.location ?? "" : link.target ?? "");
  const payload = [
    ...u323(row),
    ...u323(row),
    ...u323(col),
    ...u323(col),
    ...wideString2(relationship),
    ...wideString2(location),
    ...wideString2(link.tooltip ?? ""),
    ...wideString2(display)
  ];
  return buildBrtRecord(BRT_HLINK, Uint8Array.from(payload));
}
function literalListFormula(values) {
  return `"${values.map((value) => value.replace(/"/g, '""')).join(",")}"`;
}
function validationFormula(text) {
  if (text === void 0 || text === "") return new Uint8Array(8);
  const tokens = compileXlsbFormula(text);
  return Uint8Array.from([...u323(tokens.length), ...tokens, ...u323(0)]);
}
function buildXlsbValidationRecords(validations) {
  if (validations.length === 0) return new Uint8Array(0);
  const parts = [
    buildBrtRecord(BRT_BEGIN_DVALS, Uint8Array.from([
      ...u163(0),
      ...u323(0),
      ...u323(0),
      ...u323(0),
      ...u323(validations.length)
    ]))
  ];
  for (const rule of validations) {
    const ranges = rule.sqref.trim().split(/\s+/).map(parseXlsbRange);
    if (ranges.length === 0) throw new Error("Data validation sqref must not be empty");
    let flags = DV_TYPES2[rule.type] ?? 0;
    flags |= (DV_ERROR_STYLES[rule.errorStyle ?? "stop"] ?? 0) << 4;
    const isLiteralList = rule.type === "list" && rule.values !== void 0;
    if (isLiteralList) flags |= 128;
    if (rule.allowBlank) flags |= 256;
    if (rule.showDropdown === false) flags |= 512;
    if (rule.showInputMessage) flags |= 1 << 18;
    if (rule.showErrorMessage) flags |= 1 << 19;
    if (rule.operator !== void 0) flags |= (DV_OPERATORS2[rule.operator] ?? 0) << 20;
    const payload = [...u323(flags), ...u323(ranges.length)];
    for (const range of ranges) {
      payload.push(...u323(range.r1), ...u323(range.r2), ...u323(range.c1), ...u323(range.c2));
    }
    payload.push(
      ...nullableString(rule.errorTitle),
      ...nullableString(rule.error),
      ...nullableString(rule.promptTitle),
      ...nullableString(rule.prompt)
    );
    const formula1 = isLiteralList ? literalListFormula(rule.values) : rule.formula1;
    payload.push(...validationFormula(formula1));
    payload.push(...validationFormula(rule.formula2));
    parts.push(buildBrtRecord(BRT_DVAL, Uint8Array.from(payload)));
  }
  parts.push(buildBrtRecord(BRT_END_DVALS, new Uint8Array(0)));
  return concat(parts);
}

// src/browser/BrowserXlsbWriter.ts
var INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;
var BrowserXlsbWriter = class {
  constructor(options = {}) {
    this._useSharedStrings = options.useSharedStrings !== false;
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
    this._streamPending = null;
    this._streamHeadFlushed = true;
    this._streamHasHeaders = false;
    this._streamDataRows = 0;
    this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    this._styles = new XlsbStyleRegistry();
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
    this._stylesBin = new Uint8Array(XLSB_STYLES_TEMPLATE);
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
  _buildStylesBin() {
    return new Uint8Array(this._styles.toStylesBin(this._stylesBin));
  }
  /** Resolve column widths; explicit widths win over the conservative autofit. */
  _resolveColumnWidths(columnCount, headers, rows, options, doAutofilter) {
    const widths = initColWidths(columnCount);
    if (options.columnWidths) {
      for (let i = 0; i < columnCount; i++) {
        const width = options.columnWidths[i];
        if (width !== null && width !== void 0) widths[i] = assertColumnWidth(width, i);
      }
      return widths;
    }
    if (options.autofit === false) return widths;
    if (headers) {
      applyHeaderWidths(widths, headers, columnCount, {
        bold: options.headerStyle !== "fill",
        filterButton: doAutofilter
      });
    }
    const samples = options.sampleRows ?? rows;
    const limit = options.sampleRows ? samples?.length ?? 0 : resolveAutofitRows(options.autofitRows);
    if (samples && limit > 0) updateColWidthsFromRows(widths, samples, limit);
    return widths;
  }
  /** Store supported sheet features; reject ones XLSB does not emit yet. */
  _applySheetFeatures(sheet, features) {
    if (features.conditionalFormats && features.conditionalFormats.length > 0) {
      throw new Error("Conditional formatting is not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.");
    }
    if (features.comments && features.comments.length > 0) {
      throw new Error("Cell comments are not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.");
    }
    if (features.merges) sheet.merges = features.merges;
    if (features.dataValidations) sheet.dataValidations = features.dataValidations;
    if (features.hyperlinks && features.hyperlinks.length > 0) {
      let nextRelId = 2;
      sheet.hyperlinks = features.hyperlinks.map((spec) => {
        const range = parseXlsbRange(spec.ref);
        const internal = spec.target === void 0;
        return {
          spec,
          rId: internal ? void 0 : `rId${nextRelId++}`,
          row: range.r1,
          col: range.c1
        };
      });
    }
  }
  _writeFeatureRecords(bb, sheet) {
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
    const sheet = this._sheetList[this._sheetCount - 1];
    this._applySheetFeatures(sheet, options);
    this._isStreaming = true;
    this._currentSheetBuffer = new BrowserBigBuffer();
    this._currentSheetRowNum = 0;
    this._currentSheetStartCol = 0;
    this._currentSheetEndCol = columnCount;
    this._currentSheetDoAutofilter = doAutofilter && headers !== void 0;
    this._colWidths = this._resolveColumnWidths(
      columnCount,
      headers,
      void 0,
      options,
      doAutofilter && headers !== void 0
    );
    this._streamPending = new BrowserBigBuffer();
    this._streamHeadFlushed = false;
    this._streamHasHeaders = headers !== void 0;
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
  _flushStreamHead() {
    if (this._streamHeadFlushed) return;
    const bb = this._currentSheetBuffer;
    if (!this._isStreaming || bb === null) throw new Error("Not streaming");
    this._writeSheetHeader(bb, this._currentSheetStartCol, this._currentSheetEndCol, this._currentSheetDoAutofilter);
    const pending = this._streamPending;
    if (pending) {
      for (const chunk of pending.getChunks()) bb.write(chunk);
    }
    this._streamPending = null;
    this._streamHeadFlushed = true;
  }
  _writeStreamRow(bb, row) {
    assertRowCount(this._currentSheetRowNum + 1);
    this.createRowHeader(bb, this._currentSheetRowNum, this._currentSheetStartCol, this._currentSheetEndCol);
    for (let c = 0; c < row.length; c++) this._writeCell(bb, row[c], c);
    this._currentSheetRowNum++;
  }
  writeRow(row) {
    const bb = this._currentSheetBuffer;
    if (!this._isStreaming || bb === null) throw new Error("Not streaming");
    if (row.length !== this._currentSheetEndCol - this._currentSheetStartCol) {
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
  endSheet() {
    this._flushStreamHead();
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
    const sheet = this._sheetList[this._sheetCount - 1];
    if (!sheet) throw new Error("Call addSheet() before writeSheet()");
    this._applySheetFeatures(sheet, options);
    this._colWidths = this._resolveColumnWidths(
      colCount,
      headers,
      rows,
      options,
      doAutofilter && headers !== null
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
        endRow: rn - 1,
        startColumn: 0,
        endColumn: colCount - 1
      };
    }
    this._writeFeatureRecords(bb, sheet);
    bb.write(this._sheet1Bytes.subarray(290));
    bb.patchInt32LE(36, Math.max(0, rn - 1));
    this._zip.addFile(this._sheetList[this._sheetCount - 1].pathInArchive, bb.toUint8Array());
  }
  _writeSheetHeader(bb, startCol, endCol, doAf) {
    const sh = new Uint8Array(this._sheet1Bytes);
    const shView = new DataView(sh.buffer);
    shView.setInt32(40, startCol, true);
    shView.setInt32(44, endCol - 1, true);
    if (this._sheetCount !== 1) sh[54] = 156;
    bb.write(sh.subarray(0, 84));
    if (doAf) bb.write(this._stickHeaderA1bytes);
    bb.write(sh.subarray(84, 159));
    bb.writeByte(134);
    bb.writeByte(3);
    bb.writeByte(0);
    for (let i = startCol; i < endCol; i++) {
      bb.writeByte(60);
      bb.writeByte(18);
      bb.writeInt32LE(i);
      bb.writeInt32LE(i);
      const w = defaultColWidth(this._colWidths[i]);
      bb.writeInt32LE(Math.max(0, Math.min(255 * 256, Math.round(w * 256))));
      bb.writeInt32LE(0);
      bb.writeByte(2);
      bb.writeByte(0);
    }
    bb.writeByte(135);
    bb.writeByte(3);
    bb.writeByte(0);
    bb.write(sh.subarray(159, 175));
    bb.write(new Uint8Array([38, 0]));
  }
  _writeCell(bb, raw, c) {
    if (raw === null || raw === void 0) return;
    const style = getCellStyle(raw);
    let styleNum = style !== null ? this._styles.register(style) : 0;
    const formula = getFormula(raw);
    if (formula !== null) {
      const rawCached = raw.value;
      let cached;
      if (rawCached instanceof Date) {
        cached = timeToOaDate(rawCached.getTime(), this._oaEpoch);
        if (styleNum === 0) styleNum = dateHasTime(rawCached) ? 2 : 1;
      } else if (isRichText(rawCached)) {
        cached = rawCached.toString();
      } else if (typeof rawCached === "bigint") {
        cached = rawCached.toString();
      } else {
        cached = rawCached;
      }
      bb.write(buildXlsbFormulaRecord(c, styleNum, formula, cached));
      return;
    }
    const val = unwrapCell(raw);
    if (val === null || val === void 0) return;
    if (isRichText(val)) {
      throw new Error("Rich text is not supported by BrowserXlsbWriter yet; use BrowserXlsxWriter.");
    }
    if (typeof val === "number") {
      if (Number.isInteger(val) && val >= this._rRkIntegerLowerLimit && val <= this._rRkIntegerUpperLimit) {
        this.writeRkNumberInteger(bb, val, c, styleNum);
      } else {
        this.writeDouble(bb, val, c, styleNum);
      }
    } else if (val instanceof ExcelError) {
      this.writeError(bb, val.code, c, styleNum);
    } else if (typeof val === "bigint") {
      this.writeString(bb, val.toString(), c, 0, styleNum);
    } else if (typeof val === "boolean") {
      this.writeBool(bb, val, c, styleNum);
    } else if (val instanceof Date) {
      const oaDate = timeToOaDate(val.getTime(), this._oaEpoch);
      const dateStyle = dateHasTime(val) ? 2 : 1;
      this.writeDouble(bb, oaDate, c, styleNum !== 0 ? styleNum : dateStyle);
    } else {
      this.writeString(bb, val.toString(), c, 0, styleNum);
    }
  }
  createRowHeader(bb, rn, sc, ec) {
    writeXlsbRowHeader(bb, rn, sc, ec);
  }
  writeRkNumberInteger(bb, val, colNum, styleNum = 0) {
    writeXlsbRkInteger(bb, val, colNum, styleNum);
  }
  writeDouble(bb, val, colNum, styleNum = 0) {
    writeXlsbDouble(bb, val, colNum, styleNum);
  }
  writeBool(bb, val, colNum, styleNum = 0) {
    writeXlsbBool(bb, val, colNum, styleNum);
  }
  writeError(bb, code, colNum, styleNum = 0) {
    writeXlsbError(bb, code, colNum, styleNum);
  }
  writeDateTime(bb, date, colNum) {
    this.writeDouble(bb, timeToOaDate(date.getTime(), this._oaEpoch), colNum, dateHasTime(date) ? 2 : 1);
  }
  writeString(bb, val, colNum, styleIndex = 0, styleOverride) {
    assertCellTextLength(val);
    const finalStyle = styleOverride !== void 0 ? styleOverride : styleIndex;
    if (!this._useSharedStrings) {
      writeXlsbInlineString(bb, val, colNum, finalStyle);
      return;
    }
    let index = this._sstDic.get(val);
    if (index === void 0) {
      index = this._sstCntUnique++;
      this._sstDic.set(val, index);
    }
    this._sstCntAll++;
    writeXlsbIsst(bb, index, colNum, finalStyle);
  }
  _saveSst() {
    const bb = new BrowserBigBuffer();
    writeXlsbSharedStrings(bb, this._sstDic.keys(), this._sstCntAll, this._sstCntUnique);
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
      ct += `<Override PartName="/${s.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>`;
      ct += `<Override PartName="/xl/worksheets/binaryIndex${s.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>`;
    }
    ct += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>`;
    this._zip.addFile("[Content_Types].xml", ct);
    let wr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const s of this._sheetList) wr += `<Relationship Id="rId${s.sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.nameInArchive}"/>`;
    wr += `<Relationship Id="rId${this._sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this._sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
</Relationships>`;
    this._zip.addFile("xl/_rels/workbook.bin.rels", wr);
    this._zip.addFile("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>`);
    for (const s of this._sheetList) {
      let ws = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex${s.sheetId}.bin"/>`;
      for (const link of s.hyperlinks ?? []) {
        if (!link.rId || link.spec.target === void 0) continue;
        ws += `<Relationship Id="${link.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXmlText(link.spec.target)}" TargetMode="External"/>`;
      }
      ws += "</Relationships>";
      this._zip.addFile(`xl/worksheets/_rels/${s.nameInArchive}.rels`, ws);
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
