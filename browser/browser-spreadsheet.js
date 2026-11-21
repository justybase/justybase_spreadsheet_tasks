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
function definedEntries(value) {
  if (!value) return [];
  return Object.entries(value).filter(([, item]) => item !== void 0);
}
function mergeCellStyles(base, extra) {
  if (!base && !extra) return null;
  const out = { ...base ?? {} };
  for (const [key, value] of definedEntries(extra)) {
    if (value === void 0) continue;
    if (key === "font" || key === "fill" || key === "border" || key === "alignment" || key === "protection") {
      const current = out[key];
      out[key] = {
        ...current ?? {},
        ...value
      };
    } else {
      out[key] = value;
    }
  }
  return out;
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

// src/writers/core/sheetLayout.ts
var MAX_EXCEL_ROW = 1048575;
var MAX_EXCEL_COL = 16383;
var CELL_RE = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/;
var OPEN_COLUMNS_RE = /^\$?([A-Za-z]{1,3})(?:\$?([1-9]\d*))?:\$?([A-Za-z]{1,3})$/;
var OPEN_ROWS_RE = /^\$?([1-9]\d*):\$?([1-9]\d*)$/;
function columnIndex(letters) {
  let value = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) throw new Error(`Invalid column '${letters}'`);
    value = value * 26 + (code - 64);
  }
  return value - 1;
}
function parseCellRef(ref) {
  const match = CELL_RE.exec(ref.trim());
  if (!match) throw new Error(`Invalid cell reference '${ref}'`);
  const col = columnIndex(match[1]);
  const row = parseInt(match[2], 10) - 1;
  if (col > MAX_EXCEL_COL || row > MAX_EXCEL_ROW) {
    throw new RangeError(`Cell reference '${ref}' is outside the Excel grid`);
  }
  return { row, col };
}
function parseRangeRef(ref) {
  const parts = ref.trim().split(":");
  if (parts.length > 2) throw new Error(`Invalid range '${ref}'`);
  const first = parseCellRef(parts[0]);
  const last = parts.length === 2 ? parseCellRef(parts[1]) : first;
  return {
    c1: Math.min(first.col, last.col),
    r1: Math.min(first.row, last.row),
    c2: Math.max(first.col, last.col),
    r2: Math.max(first.row, last.row)
  };
}
function parseStyleRange(ref) {
  const text = ref.trim();
  if (!text) throw new Error("Style range must not be empty");
  const columns = OPEN_COLUMNS_RE.exec(text);
  if (columns) {
    const c1 = columnIndex(columns[1]);
    const c2 = columnIndex(columns[3]);
    const r1 = columns[2] === void 0 ? 0 : parseInt(columns[2], 10) - 1;
    if (c1 > MAX_EXCEL_COL || c2 > MAX_EXCEL_COL || r1 > MAX_EXCEL_ROW) {
      throw new RangeError(`Style range '${ref}' is outside the Excel grid`);
    }
    return { c1: Math.min(c1, c2), r1, c2: Math.max(c1, c2), r2: MAX_EXCEL_ROW };
  }
  const rows = OPEN_ROWS_RE.exec(text);
  if (rows) {
    const r1 = parseInt(rows[1], 10) - 1;
    const r2 = parseInt(rows[2], 10) - 1;
    if (r1 > MAX_EXCEL_ROW || r2 > MAX_EXCEL_ROW) {
      throw new RangeError(`Style range '${ref}' is outside the Excel grid`);
    }
    return { c1: 0, r1: Math.min(r1, r2), c2: MAX_EXCEL_COL, r2: Math.max(r1, r2) };
  }
  return parseRangeRef(text);
}
var StyleRangeSet = class {
  constructor(styles) {
    this._row = -1;
    this._active = [];
    if (!styles) {
      this.ranges = [];
      return;
    }
    if (Array.isArray(styles)) {
      this.ranges = styles;
      return;
    }
    this.ranges = Object.entries(styles).map(([ref, style]) => {
      if (!style || typeof style !== "object") {
        throw new TypeError(`Style for '${ref}' must be a CellStyle object`);
      }
      return { ...parseStyleRange(ref), style };
    });
  }
  get length() {
    return this.ranges.length;
  }
  /** Ranges covering `rowIndex`, in the caller's order. */
  active(rowIndex) {
    if (rowIndex === this._row) return this._active;
    this._row = rowIndex;
    this._active = this.ranges.filter((range) => range.r1 <= rowIndex && rowIndex <= range.r2);
    return this._active;
  }
  /** Merge every range style covering the cell, in order. */
  styleFor(rowIndex, columnIndex4) {
    let style = null;
    for (const range of this.active(rowIndex)) {
      if (range.c1 <= columnIndex4 && columnIndex4 <= range.c2) {
        style = mergeCellStyles(style, range.style);
      }
    }
    return style;
  }
};
function normalizeGroupValue(value, kind) {
  if (typeof value === "number") {
    if (!Number.isInteger(value) || value < 0) {
      throw new RangeError(`Group ${kind} must be a non-negative integer, got ${value}`);
    }
    return value;
  }
  if (kind === "row") {
    const parsed = parseInt(value, 10);
    if (!Number.isInteger(parsed) || parsed < 1) {
      throw new RangeError(`Row group bounds must be 1-based row numbers, got '${value}'`);
    }
    return parsed - 1;
  }
  return columnIndex(value);
}
function normalizeGroups(groups, kind) {
  if (!groups) return [];
  return groups.map((group) => {
    const start = normalizeGroupValue(group.start, kind);
    const end = normalizeGroupValue(group.end, kind);
    if (end < start) throw new RangeError(`Group end precedes start in ${kind} group`);
    const limit = kind === "row" ? MAX_EXCEL_ROW : MAX_EXCEL_COL;
    if (end > limit) throw new RangeError(`${kind} group is outside the Excel grid`);
    return { start, end, collapsed: group.collapsed === true };
  });
}
function buildOutline(groups, maxIndex) {
  const size = Math.max(0, maxIndex + 1);
  const level = new Uint8Array(size);
  const hidden = new Uint8Array(size);
  const collapsed = new Uint8Array(size);
  const indices = /* @__PURE__ */ new Set();
  for (let i = 0; i < groups.length; i++) {
    for (let j = i + 1; j < groups.length; j++) {
      const a = groups[i];
      const b = groups[j];
      const same = a.start === b.start && a.end === b.end;
      if (same) throw new RangeError("Duplicate row/column group");
      const disjoint = a.end < b.start || b.end < a.start;
      const nested = a.start <= b.start && b.end <= a.end || b.start <= a.start && a.end <= b.end;
      if (!disjoint && !nested) {
        throw new RangeError("Row/column groups may nest or be disjoint, not overlap partially");
      }
    }
  }
  for (const group of groups) {
    const depth = groups.filter(
      (other) => other !== group && other.start <= group.start && group.end <= other.end
    ).length + 1;
    if (depth > 7) throw new RangeError("Excel supports at most seven outline levels");
    for (let index = group.start; index <= group.end && index < size; index++) {
      level[index] = Math.max(level[index], depth);
      hidden[index] = 1;
      indices.add(index);
    }
    const summary = group.end + 1;
    if (group.collapsed && summary < size) {
      collapsed[summary] = 1;
      indices.add(summary);
    }
  }
  let maxLevel = 0;
  for (const value of level) if (value > maxLevel) maxLevel = value;
  return { level, hidden, collapsed, maxLevel, indices: [...indices].sort((a, b) => a - b) };
}
function resolveFreezePanes(value, hasHeaders) {
  let ref;
  if (value === void 0) ref = hasHeaders ? "A2" : null;
  else if (value === false || value === null || value === "") ref = null;
  else ref = value;
  if (ref === null) return null;
  const { row, col } = parseCellRef(ref);
  if (row === 0 && col === 0) return null;
  const activePane = row > 0 && col > 0 ? "bottomRight" : col > 0 ? "topRight" : "bottomLeft";
  return {
    xSplit: col,
    ySplit: row,
    topLeftCell: `${columnIndexToLetter(col)}${row + 1}`,
    activePane,
    frozen: true
  };
}
function validateZoom(value) {
  if (value === void 0) return void 0;
  if (!Number.isInteger(value) || value < 10 || value > 400) {
    throw new RangeError(`zoom must be an integer 10..400, got ${value}`);
  }
  return value;
}
function validateRowHeight(value) {
  if (!Number.isFinite(value) || value <= 0 || value > 409.5) {
    throw new RangeError(`Row height must be in (0, 409.5] points, got ${value}`);
  }
  return value;
}
function passwordHash(password) {
  let hash = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    hash = hash >> 14 & 1 | hash << 1 & 32767;
    hash ^= password.charCodeAt(i) & 255;
  }
  hash = hash >> 14 & 1 | hash << 1 & 32767;
  hash ^= 52811;
  hash ^= password.length;
  return hash & 65535;
}
var TABLE_NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
var A1_REFERENCE_RE = /^([A-Za-z]{1,3})([1-9]\d*)$/;
var R1C1_REFERENCE_RE = /^R[1-9]\d*C[1-9]\d*$/i;
function tableNameIsCellReference(name) {
  const a1 = A1_REFERENCE_RE.exec(name);
  if (a1) {
    let column = 0;
    for (const ch of a1[1].toUpperCase()) column = column * 26 + (ch.charCodeAt(0) - 64);
    const row = Number(a1[2]);
    if (column <= 16384 && row <= 1048576) return true;
  }
  if (R1C1_REFERENCE_RE.test(name)) return true;
  return /^[RC]$/i.test(name);
}
function normalizeTableSpec(spec, id) {
  if (!spec || typeof spec !== "object") throw new TypeError("tables must be TableSpec objects");
  const name = (spec.name ?? `Table${id}`).trim();
  if (!TABLE_NAME_RE.test(name)) {
    throw new Error(`Invalid table name '${name}'`);
  }
  if (name.length > 255) {
    throw new Error(`Table name '${name}' exceeds 255 characters`);
  }
  if (tableNameIsCellReference(name)) {
    throw new Error(`Invalid table name '${name}': it is a cell reference`);
  }
  const ref = parseRangeRef(spec.ref);
  const width = ref.c2 - ref.c1 + 1;
  const columns = (spec.columns ?? []).map((column) => {
    if (typeof column === "string") return { name: column };
    if (!column || typeof column.name !== "string" || !column.name) {
      throw new Error(`Table '${name}' has a column without a name`);
    }
    return { ...column };
  });
  if (columns.length === 0) {
    throw new Error(`Table '${name}' needs columns (pass TableSpec.columns)`);
  }
  if (columns.length !== width) {
    throw new Error(
      `Table '${name}' has ${columns.length} columns but ref '${spec.ref}' is ${width} columns wide`
    );
  }
  const names = /* @__PURE__ */ new Set();
  for (const column of columns) {
    if (names.has(column.name.toLowerCase())) {
      throw new Error(`Table '${name}' has duplicate column '${column.name}'`);
    }
    names.add(column.name.toLowerCase());
  }
  return {
    name,
    ref,
    refText: `${columnIndexToLetter(ref.c1)}${ref.r1 + 1}:${columnIndexToLetter(ref.c2)}${ref.r2 + 1}`,
    columns,
    headerRow: spec.headerRow !== false,
    totalsRow: spec.totalsRow === true,
    styleName: spec.styleName ?? "TableStyleMedium9",
    showFirstColumn: spec.showFirstColumn === true,
    showLastColumn: spec.showLastColumn === true,
    showRowStripes: spec.showRowStripes !== false,
    showColumnStripes: spec.showColumnStripes === true
  };
}
function tableAutoFilterRef(table) {
  if (!table.headerRow) return null;
  const endRow = table.totalsRow ? table.ref.r2 - 1 : table.ref.r2;
  if (endRow < table.ref.r1) return null;
  return `${columnIndexToLetter(table.ref.c1)}${table.ref.r1 + 1}:${columnIndexToLetter(table.ref.c2)}${endRow + 1}`;
}
var TOTALS_SUBTOTAL_CODE = {
  average: 101,
  count: 103,
  countNums: 102,
  max: 104,
  min: 105,
  sum: 109,
  stdDev: 107,
  var: 108
};
function tableTotalsOverrides(tables) {
  const overrides = [];
  for (const table of tables) {
    if (!table.totalsRow) continue;
    for (let index = 0; index < table.columns.length; index++) {
      const column = table.columns[index];
      const col = table.ref.c1 + index;
      const fn = column.totalsRowFunction;
      if (fn === "custom") {
        const formula = column.totalsRowFormula?.trim().replace(/^=/, "");
        if (!formula) {
          throw new Error(
            `Table '${table.name}' column '${column.name}' needs totalsRowFormula for a custom total`
          );
        }
        overrides.push({ row: table.ref.r2, col, formula });
      } else if (fn !== void 0) {
        const code = TOTALS_SUBTOTAL_CODE[fn];
        const firstDataRow = table.headerRow ? table.ref.r1 + 1 : table.ref.r1;
        const lastDataRow = table.ref.r2 - 1;
        const letter = columnIndexToLetter(col);
        if (lastDataRow < firstDataRow) {
          throw new Error(
            `Table '${table.name}' needs at least one data row before its totals row`
          );
        }
        overrides.push({
          row: table.ref.r2,
          col,
          formula: `SUBTOTAL(${code},${letter}${firstDataRow + 1}:${letter}${lastDataRow + 1})`
        });
      } else if (column.totalsRowLabel !== void 0) {
        overrides.push({ row: table.ref.r2, col, text: column.totalsRowLabel });
      }
    }
  }
  return overrides;
}
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
  if ("auto" in color) return "auto";
  return `theme:${color.theme}:${color.tint ?? 0}`;
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
function colorKey2(color) {
  return colorKey(color);
}
function colorXml(color) {
  if (color === void 0) return "";
  if (typeof color === "string") {
    const key = colorKey2(color);
    return `<color rgb="${key.slice(4)}"/>`;
  }
  if ("auto" in color) return '<color auto="1"/>';
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
  if (fill.gradient !== void 0) {
    const gradient = fill.gradient;
    const type = gradient.type ?? "linear";
    if (type !== "linear" && type !== "path") {
      throw new Error(`Unknown gradient fill type '${type}'`);
    }
    if (!Array.isArray(gradient.stops) || gradient.stops.length < 2) {
      throw new Error("A gradient fill needs at least two stops");
    }
    const stops = gradient.stops.map((stop) => {
      if (!Number.isFinite(stop.position) || stop.position < 0 || stop.position > 1) {
        throw new RangeError(`Gradient stop position must be 0..1, got ${stop.position}`);
      }
      return { color: stop.color, position: stop.position };
    });
    if (gradient.degree !== void 0 && (!Number.isFinite(gradient.degree) || gradient.degree < 0 || gradient.degree > 360)) {
      throw new RangeError(`Gradient degree must be 0..360, got ${gradient.degree}`);
    }
    for (const side of ["left", "right", "top", "bottom"]) {
      const value = gradient[side];
      if (value !== void 0 && (!Number.isFinite(value) || value < 0 || value > 1)) {
        throw new RangeError(`Gradient ${side} must be 0..1, got ${value}`);
      }
    }
    return { gradient: { ...gradient, type, stops } };
  }
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
  for (const name of ["top", "bottom", "left", "right", "diagonal", "start", "end", "vertical", "horizontal"]) {
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
  if (alignment.relativeIndent !== void 0) {
    if (!Number.isInteger(alignment.relativeIndent) || alignment.relativeIndent < -250 || alignment.relativeIndent > 250) {
      throw new RangeError(`alignment.relativeIndent must be -250..250, got ${alignment.relativeIndent}`);
    }
    out.relativeIndent = alignment.relativeIndent;
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
  if (style.quotePrefix) out.quotePrefix = true;
  if (style.pivotButton) out.pivotButton = true;
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
  if (font.bold) xml += "<b/>";
  if (font.italic) xml += "<i/>";
  if (font.strike) xml += "<strike/>";
  if (font.condense) xml += "<condense/>";
  if (font.extend) xml += "<extend/>";
  if (font.outline) xml += "<outline/>";
  if (font.shadow) xml += "<shadow/>";
  if (font.underline !== void 0) {
    xml += font.underline === "single" ? "<u/>" : `<u val="${font.underline}"/>`;
  }
  if (font.vertAlign !== void 0) xml += `<vertAlign val="${font.vertAlign}"/>`;
  if (font.size !== void 0) xml += `<sz val="${font.size}"/>`;
  if (font.color !== void 0) xml += colorXml(font.color);
  if (font.name !== void 0) xml += `<name val="${escapeXlsxAttribute(font.name)}"/>`;
  if (font.family !== void 0) xml += `<family val="${font.family}"/>`;
  if (font.charset !== void 0) xml += `<charset val="${font.charset}"/>`;
  if (font.scheme !== void 0) xml += `<scheme val="${font.scheme}"/>`;
  return `${xml}</font>`;
}
function fillXml(fill) {
  if (fill.gradient) {
    const gradient = fill.gradient;
    let attrs = ` type="${gradient.type ?? "linear"}"`;
    if (gradient.degree !== void 0) attrs += ` degree="${gradient.degree}"`;
    for (const side of ["left", "right", "top", "bottom"]) {
      if (gradient[side] !== void 0) attrs += ` ${side}="${gradient[side]}"`;
    }
    let xml2 = `<fill><gradientFill${attrs}>`;
    for (const stop of gradient.stops) {
      xml2 += `<stop position="${stop.position}">${colorXml(stop.color)}</stop>`;
    }
    return `${xml2}</gradientFill></fill>`;
  }
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
  if (border.start !== void 0) xml += borderSideXml("start", border.start);
  if (border.end !== void 0) xml += borderSideXml("end", border.end);
  xml += borderSideXml("left", border.left);
  xml += borderSideXml("right", border.right);
  xml += borderSideXml("top", border.top);
  xml += borderSideXml("bottom", border.bottom);
  xml += borderSideXml("diagonal", border.diagonal);
  if (border.vertical !== void 0) xml += borderSideXml("vertical", border.vertical);
  if (border.horizontal !== void 0) xml += borderSideXml("horizontal", border.horizontal);
  return `${xml}</border>`;
}
function alignmentXml(alignment) {
  let xml = "<alignment";
  if (alignment.horizontal !== void 0) xml += ` horizontal="${alignment.horizontal}"`;
  if (alignment.vertical !== void 0) xml += ` vertical="${alignment.vertical}"`;
  if (alignment.wrapText) xml += ' wrapText="1"';
  if (alignment.shrinkToFit) xml += ' shrinkToFit="1"';
  if (alignment.indent !== void 0) xml += ` indent="${alignment.indent}"`;
  if (alignment.relativeIndent !== void 0) xml += ` relativeIndent="${alignment.relativeIndent}"`;
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
function dxfFillXml(fill) {
  if (fill.gradient) return fillXml(fill);
  const pattern = fill.patternType ?? "solid";
  if (pattern !== "solid") return fillXml(fill);
  const color = fill.bgColor ?? fill.fgColor;
  if (color === void 0) return fillXml(fill);
  return `<fill><patternFill patternType="solid">${colorXml(color).replace("<color", "<bgColor")}</patternFill></fill>`;
}
function dxfXml(style, numFmtId) {
  let xml = "<dxf>";
  if (style.font) xml += fontXml(style.font);
  if (style.numFmt !== void 0 && numFmtId !== null) {
    xml += `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXlsxAttribute(style.numFmt)}"/>`;
  }
  if (style.fill) {
    xml += dxfFillXml(style.fill);
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
    if (normalized.quotePrefix) xf += ' quotePrefix="1"';
    if (normalized.pivotButton) xf += ' pivotButton="1"';
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
function buildXlsxContentTypes(sheets, comments, tables) {
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
  for (const part of tables ?? []) {
    xml += `<Override PartName="/${part.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`;
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
function xlsxRowAttributes(rowIndex, layout) {
  if (!layout) return "";
  let attrs = "";
  const height = layout.rowHeights?.[rowIndex];
  if (height !== void 0) {
    attrs += ` ht="${height}" customHeight="1"`;
  }
  const outline = layout.rowOutline;
  const outlineHidden = outline !== void 0 && rowIndex < outline.hidden.length && outline.hidden[rowIndex] === 1;
  if (layout.hiddenRows?.has(rowIndex) || outlineHidden) attrs += ' hidden="1"';
  if (outline && rowIndex < outline.level.length) {
    if (outline.level[rowIndex] > 0) attrs += ` outlineLevel="${outline.level[rowIndex]}"`;
    if (outline.collapsed[rowIndex]) attrs += ' collapsed="1"';
  }
  return attrs;
}
function writeXlsxSheetHead(sink, options) {
  const {
    dimensionRef,
    columnCount,
    isFirstSheet,
    colWidths,
    freezePanes = null,
    showGridLines = true,
    showRowColHeaders = true,
    showZeroValues = true,
    rightToLeft = false,
    zoom,
    tabColor,
    hiddenColumns = [],
    columnOutline
  } = options;
  sink.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  sink.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
  if (tabColor !== void 0) {
    sink.writeString(`<sheetPr><tabColor ${colorXml(tabColor).replace(/^<color /, "").replace(/\/>$/, "")}/></sheetPr>`);
  }
  sink.writeString(`<dimension ref="${dimensionRef}"/>`);
  let viewAttrs = "";
  if (isFirstSheet) viewAttrs += ' tabSelected="1"';
  if (!showGridLines) viewAttrs += ' showGridLines="0"';
  if (!showRowColHeaders) viewAttrs += ' showRowColHeaders="0"';
  if (!showZeroValues) viewAttrs += ' showZeros="0"';
  if (rightToLeft) viewAttrs += ' rightToLeft="1"';
  if (zoom !== void 0) viewAttrs += ` zoomScale="${validateZoom(zoom)}"`;
  viewAttrs += ' workbookViewId="0"';
  sink.writeString("<sheetViews>");
  if (freezePanes) {
    const pane = freezePanes;
    sink.writeString(`<sheetView${viewAttrs}>`);
    sink.writeString(`<pane${pane.xSplit ? ` xSplit="${pane.xSplit}"` : ""}${pane.ySplit ? ` ySplit="${pane.ySplit}"` : ""} topLeftCell="${pane.topLeftCell}" activePane="${pane.activePane}" state="frozen"/>`);
    sink.writeString(`<selection pane="${pane.activePane}" activeCell="${pane.topLeftCell}" sqref="${pane.topLeftCell}"/>`);
    sink.writeString("</sheetView>");
  } else {
    sink.writeString(`<sheetView${viewAttrs}/>`);
  }
  sink.writeString("</sheetViews>");
  let formatPr = '<sheetFormatPr defaultRowHeight="15"';
  if (columnOutline && columnOutline.maxLevel > 0) formatPr += ` outlineLevelCol="${columnOutline.maxLevel}"`;
  formatPr += "/>";
  sink.writeString(formatPr);
  sink.writeString("<cols>");
  const hidden = new Set(hiddenColumns);
  for (let i = 0; i < columnCount; i++) {
    const width = defaultColWidth(colWidths[i]);
    let attrs = `min="${i + 1}" max="${i + 1}" width="${formatColumnWidth(width)}" bestFit="1" customWidth="1"`;
    if (hidden.has(i)) attrs += ' hidden="1"';
    if (columnOutline && i < columnOutline.level.length && columnOutline.level[i] > 0) {
      attrs += ` outlineLevel="${columnOutline.level[i]}"`;
      if (columnOutline.collapsed[i]) attrs += ' collapsed="1"';
    }
    sink.writeString(`<col ${attrs}/>`);
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
    registerDxf,
    sheetProtection,
    autoFilterRef,
    tableParts
  } = options;
  sink.writeString("</sheetData>");
  if (sheetProtection) sink.writeString(buildXlsxSheetProtection(sheetProtection));
  let filterRange = null;
  if (doAutofilter && columnCount > 0) {
    let ref = `A1:${lastColumnLetter}${rowNum}`;
    if (autoFilterRef !== void 0) ref = autoFilterRef;
    sink.writeString(`<autoFilter ref="${escapeXmlText(ref)}"/>`);
    const name = autoFilterRef !== void 0 ? autoFilterRef.replace(/\$?([A-Za-z]{1,3})\$?([0-9]+)/g, "$1$2") : `A1:${lastColumnLetter}${rowNum}`;
    const absolute = name.replace(/^([A-Za-z]+)(\d+)/, "$$$1$$$2").replace(/:([A-Za-z]+)(\d+)$/, ":$$$1$$$2");
    filterRange = `${formatSheetNameForFormula(sheetName)}!${absolute}`;
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
  if (tableParts && tableParts.length > 0) {
    sink.writeString(`<tableParts count="${tableParts.length}">`);
    for (const part of tableParts) sink.writeString(`<tablePart r:id="${part.rId}"/>`);
    sink.writeString("</tableParts>");
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
var XLSX_REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
function buildXlsxTableXml(table, tableId) {
  const totals = table.totalsRow ? ' totalsRowCount="1" totalsRowShown="1"' : "";
  const headerCount = table.headerRow ? "1" : "0";
  const filterRef = tableAutoFilterRef(table);
  const auto = filterRef ? `<autoFilter ref="${escapeXmlText(filterRef)}"/>` : "";
  const columns = table.columns.map((column, index) => {
    let attrs = `id="${index + 1}" name="${escapeXlsxAttribute(column.name)}"`;
    let child = "";
    if (table.totalsRow) {
      const fn = column.totalsRowFunction;
      if (fn === "custom") {
        attrs += ' totalsRowFunction="custom"';
        child = `<totalsRowFormula>${escapeXmlText(column.totalsRowFormula ?? "")}</totalsRowFormula>`;
      } else if (fn !== void 0) {
        attrs += ` totalsRowFunction="${fn}"`;
      } else if (column.totalsRowLabel !== void 0) {
        attrs += ` totalsRowLabel="${escapeXlsxAttribute(column.totalsRowLabel)}"`;
      }
    }
    return child ? `<tableColumn ${attrs}>${child}</tableColumn>` : `<tableColumn ${attrs}/>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tableId}" name="${escapeXlsxAttribute(table.name)}" displayName="${escapeXlsxAttribute(table.name)}" ref="${table.refText}"${totals} headerRowCount="${headerCount}">${auto}<tableColumns count="${table.columns.length}">${columns}</tableColumns><tableStyleInfo name="${escapeXlsxAttribute(table.styleName)}" showFirstColumn="${table.showFirstColumn ? 1 : 0}" showLastColumn="${table.showLastColumn ? 1 : 0}" showRowStripes="${table.showRowStripes ? 1 : 0}" showColumnStripes="${table.showColumnStripes ? 1 : 0}"/></table>`;
}
function buildXlsxSheetProtection(protection) {
  if (protection.enabled === false) return "";
  const selection = (value) => value === false ? "1" : "0";
  let xml = `<sheetProtection sheet="1" objects="${protection.objects === false ? 0 : 1}" scenarios="${protection.scenarios === false ? 0 : 1}"`;
  if (protection.password) {
    xml += ` password="${passwordHash(protection.password).toString(16).toUpperCase().padStart(4, "0")}"`;
  }
  if (protection.selectLockedCells !== void 0) {
    xml += ` selectLockedCells="${selection(protection.selectLockedCells)}"`;
  }
  if (protection.selectUnlockedCells !== void 0) {
    xml += ` selectUnlockedCells="${selection(protection.selectUnlockedCells)}"`;
  }
  const locks = [
    ["formatCells", protection.formatCells],
    ["formatColumns", protection.formatColumns],
    ["formatRows", protection.formatRows],
    ["insertColumns", protection.insertColumns],
    ["insertRows", protection.insertRows],
    ["insertHyperlinks", protection.insertHyperlinks],
    ["deleteColumns", protection.deleteColumns],
    ["deleteRows", protection.deleteRows],
    ["sort", protection.sort],
    ["autoFilter", protection.autoFilter],
    ["pivotTables", protection.pivotTables]
  ];
  for (const [name, value] of locks) {
    if (value !== void 0) xml += ` ${name}="${value ? 1 : 0}"`;
  }
  return `${xml}/>`;
}
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
            sink.writeString(`<color rgb="${colorKey2(color).slice(4)}"/>`);
          }
          sink.writeString("</colorScale></cfRule>");
          break;
        }
        case "dataBar": {
          sink.writeString(`<cfRule type="dataBar" priority="${current}"><dataBar${rule.showValue === false ? ' showValue="0"' : ""}>`);
          for (const point of rule.cfvo) {
            sink.writeString(`<cfvo type="${point.type}"${point.value !== void 0 ? ` val="${escapeXmlText(point.value)}"` : ""}/>`);
          }
          sink.writeString(`<color rgb="${colorKey2(rule.color).slice(4)}"/></dataBar></cfRule>`);
          break;
        }
        case "iconSet": {
          const iconSet = rule.iconSet ?? "3TrafficLights1";
          const iconCount = iconSet.startsWith("5") ? 5 : iconSet.startsWith("4") ? 4 : 3;
          const defaults = Array.from({ length: iconCount - 1 }, (_, index) => ({
            type: "percent",
            value: String(Math.round((index + 1) * 100 / iconCount))
          }));
          const supplied = rule.cfvo ?? [];
          const points = supplied.length >= iconCount - 1 ? supplied.slice(0, iconCount - 1) : [...supplied, ...defaults.slice(supplied.length)];
          let iconAttrs = ` iconSet="${iconSet}"`;
          if (rule.showValue === false) iconAttrs += ' showValue="0"';
          if (rule.reverse) iconAttrs += ' reverse="1"';
          sink.writeString(`<cfRule type="iconSet" priority="${current}"><iconSet${iconAttrs}>`);
          for (const point of points) {
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
          const quoted = `"${escaped}"`;
          const formula = rule.type === "containsText" ? `NOT(ISERROR(SEARCH(${quoted},${firstCell})))` : rule.type === "notContainsText" ? `ISERROR(SEARCH(${quoted},${firstCell}))` : rule.type === "beginsWith" ? `LEFT(${firstCell},${rule.text.length})=${quoted}` : `RIGHT(${firstCell},${rule.text.length})=${quoted}`;
          sink.writeString(`<cfRule type="${rule.type}" operator="${operator}" text="${escapeXlsxAttribute(rule.text)}" priority="${current}"${dxfAttr}><formula>${escapeXmlText(formula)}</formula></cfRule>`);
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
function writeXlsxCell(sink, context, raw, col, prevCol, colLetter, rowNum, dateStyle = 1, dateTimeStyle = 2, extraStyle) {
  if (raw === null || raw === void 0) return prevCol;
  const cellStyle = getCellStyle(raw);
  const style = extraStyle ? mergeCellStyles(cellStyle, extraStyle) : cellStyle;
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
function withTableTotals(row, overrides) {
  let result = null;
  for (const override of overrides) {
    if (override.col < 0) continue;
    const original = row[override.col];
    const record = original !== null && typeof original === "object" && !(original instanceof Date) ? original : null;
    const style = record && record.style ? record.style : void 0;
    const cachedValue = record?.value;
    const cached = typeof cachedValue === "string" || typeof cachedValue === "number" || typeof cachedValue === "boolean" || cachedValue instanceof Date ? cachedValue : typeof original === "number" ? original : void 0;
    const replacement = override.text !== void 0 ? style ? { value: override.text, style } : override.text : {
      formula: override.formula,
      ...cached !== void 0 ? { value: cached } : {},
      ...style ? { style } : {}
    };
    if (!result) result = row.slice();
    result[override.col] = replacement;
  }
  return result ?? row;
}
var BrowserXlsxWriter = class {
  constructor(options = {}) {
    this._sheetCount = 0;
    this._sheets = [];
    this._sstValues = [];
    this._sstIndex = /* @__PURE__ */ new Map();
    this._sstCount = 0;
    this._autofilterOn = false;
    this._nextTableId = 1;
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
    this._streamTableTotals = /* @__PURE__ */ new Map();
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
  _resolveColumnWidths(columnCount, headers, rows, options, doAutofilter, headerStyle) {
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
        bold: headerStyle ? headerStyle.font?.bold === true : true,
        filterButton: doAutofilter
      });
    }
    const samples = options.sampleRows ?? rows;
    const limit = options.sampleRows ? samples?.length ?? 0 : resolveAutofitRows(options.autofitRows);
    if (samples && limit > 0) updateColWidthsFromRows(widths, samples, limit);
    return widths;
  }
  _presetHeaderStyle(options) {
    if (options.headerStyle && typeof options.headerStyle !== "string") {
      return mergeCellStyles({ font: { bold: true } }, options.headerStyle);
    }
    const preset = options.headerPreset ?? (typeof options.headerStyle === "string" ? options.headerStyle : void 0) ?? "bold";
    if (preset === "bold") return null;
    const base = {};
    if (preset === "bold+fill") base.font = { bold: true };
    if (preset === "fill" || preset === "bold+fill") {
      base.fill = { patternType: "solid", fgColor: "FFD9E1F2" };
    }
    return base;
  }
  _headerStyleIndex(layout) {
    return layout.headerStyle ? this._styles.register(layout.headerStyle) : 3;
  }
  _resolveLayout(options, hasHeaders, columnCount, headers, rowCount = 0) {
    const rowHeights = {};
    for (const [key, value] of Object.entries(options.rowHeights ?? {})) {
      const row = Number(key);
      if (!Number.isInteger(row) || row < 0) {
        throw new RangeError(`rowHeights keys must be non-negative row indices, got '${key}'`);
      }
      rowHeights[row] = validateRowHeight(value);
    }
    const hiddenRows = /* @__PURE__ */ new Set();
    for (const row of options.hiddenRows ?? []) {
      if (!Number.isInteger(row) || row < 0) {
        throw new RangeError(`hiddenRows must contain non-negative row indices, got ${row}`);
      }
      hiddenRows.add(row);
    }
    const hiddenColumns = [];
    for (const col of options.hiddenColumns ?? []) {
      if (!Number.isInteger(col) || col < 0 || col >= columnCount) {
        throw new RangeError(`hiddenColumns index ${col} is outside the sheet`);
      }
      hiddenColumns.push(col);
    }
    const rowGroups = normalizeGroups(options.rowGroups, "row");
    const columnGroups = normalizeGroups(options.columnGroups, "column");
    const rowMax = Math.max(
      rowCount - 1,
      ...rowGroups.map((group) => group.end),
      ...Object.keys(rowHeights).map(Number),
      ...hiddenRows,
      0
    );
    const columnMax = Math.max(
      columnCount - 1,
      ...columnGroups.map((group) => group.end),
      ...hiddenColumns,
      0
    );
    const rowOutline = buildOutline(rowGroups, rowMax);
    const columnOutline = buildOutline(columnGroups, columnMax);
    const tables = [];
    for (const spec of options.tables ?? []) {
      let columns = spec.columns;
      if ((!columns || columns.length === 0) && headers) {
        const range = parseRangeRef(spec.ref);
        columns = headers.slice(range.c1, range.c2 + 1);
      }
      tables.push(normalizeTableSpec({ ...spec, columns }, this._nextTableId++));
    }
    if (options.autoFilterRef !== void 0) parseRangeRef(options.autoFilterRef);
    return {
      headerStyle: this._presetHeaderStyle(options),
      styleRanges: new StyleRangeSet(options.styles),
      rowHeights,
      hiddenRows,
      hiddenColumns,
      rowOutline,
      columnOutline,
      freezePanes: resolveFreezePanes(options.freezePanes, hasHeaders),
      showGridLines: options.showGridLines !== false,
      showRowColHeaders: options.showRowColHeaders !== false,
      showZeroValues: options.showZeroValues !== false,
      rightToLeft: options.rightToLeft === true,
      zoom: validateZoom(options.zoom),
      tabColor: options.tabColor,
      sheetProtection: options.sheetProtection,
      autoFilterRef: options.autoFilterRef,
      tables
    };
  }
  _buildTableParts(sheet, layout, nextRelId) {
    if (layout.tables.length === 0) return nextRelId;
    const parts = [];
    for (const table of layout.tables) {
      const id = this._nextTableId - layout.tables.length + parts.length;
      const path = `xl/tables/table${id}.xml`;
      parts.push({
        id,
        path,
        xml: buildXlsxTableXml(table, id),
        rId: `rId${nextRelId++}`
      });
    }
    sheet.tableParts = parts;
    sheet.extraParts = [...sheet.extraParts ?? [], ...parts.map((part) => ({ name: part.path, data: part.xml }))];
    return nextRelId;
  }
  /** Store sheet features and build the worksheet rels/comments parts. */
  _applySheetFeatures(sheet, features, layout) {
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
    if (layout && layout.tables.length > 0) {
      nextRelId = this._buildTableParts(sheet, layout, nextRelId);
      for (const part of sheet.tableParts ?? []) {
        rels.push({
          id: part.rId,
          type: XLSX_REL_TABLE,
          target: `../tables/${part.path.split("/").pop()}`
        });
      }
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
    const { hidden = false, doAutofilter = true } = options;
    this.addSheet(sheetName, hidden);
    const sheet = this._sheets[this._sheetCount - 1];
    const layout = this._resolveLayout(options, headers !== void 0, columnCount, headers);
    sheet.layout = layout;
    this._streamTableTotals = /* @__PURE__ */ new Map();
    for (const override of tableTotalsOverrides(layout.tables)) {
      const list = this._streamTableTotals.get(override.row) ?? [];
      list.push(override);
      this._streamTableTotals.set(override.row, list);
    }
    this._applySheetFeatures(sheet, options, layout);
    this._streamBuffer = new BrowserBigBuffer();
    this._streamRowNum = 0;
    this._streamColumnCount = columnCount;
    this._streamDoAutofilter = doAutofilter && headers !== void 0;
    this._streamColWidths = this._resolveColumnWidths(
      columnCount,
      headers,
      void 0,
      options,
      doAutofilter && headers !== void 0,
      layout.headerStyle ?? void 0
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
      pending.writeString(`<row r="${this._streamRowNum}"${xlsxRowAttributes(0, layout)}>`);
      const headerStyle = this._headerStyleIndex(layout);
      const baseHeader = layout.headerStyle ?? { font: { bold: true } };
      for (let c = 0; c < headers.length; c++) {
        const rangeStyle = layout.styleRanges.styleFor(0, c);
        const style = rangeStyle ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle)) : headerStyle;
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
    const layout = this._sheets[this._sheetCount - 1]?.layout;
    writeXlsxSheetHead(buffer, {
      dimensionRef: "A1",
      columnCount: this._streamColumnCount,
      isFirstSheet: this._sheetCount === 1,
      hasHeaders: this._streamHasHeaders,
      colWidths: this._streamColWidths,
      freezePanes: layout?.freezePanes,
      showGridLines: layout?.showGridLines,
      showRowColHeaders: layout?.showRowColHeaders,
      showZeroValues: layout?.showZeroValues,
      rightToLeft: layout?.rightToLeft,
      zoom: layout?.zoom,
      tabColor: layout?.tabColor,
      hiddenColumns: layout?.hiddenColumns,
      columnOutline: layout?.columnOutline
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
    const layout = this._sheets[this._sheetCount - 1]?.layout;
    const totals = this._streamTableTotals.get(this._streamRowNum - 1);
    if (totals) row = withTableTotals(row, totals);
    buffer.writeString(`<row r="${this._streamRowNum}"${xlsxRowAttributes(this._streamRowNum - 1, layout)}>`);
    let prevCol = -1;
    for (let c = 0; c < row.length; c++) {
      const rangeStyle = layout?.styleRanges.styleFor(this._streamRowNum - 1, c);
      prevCol = writeXlsxCell(
        buffer,
        this,
        row[c],
        c,
        prevCol,
        this._streamColLetters[c],
        this._streamRowNum,
        1,
        2,
        rangeStyle
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
      registerDxf: (style) => this._styles.registerDxf(style),
      sheetProtection: sheet.layout?.sheetProtection,
      autoFilterRef: sheet.layout?.autoFilterRef,
      tableParts: sheet.tableParts
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
    this._streamTableTotals = /* @__PURE__ */ new Map();
  }
  // Batch API ---------------------------------------------------------
  writeSheet(rows, headers = null, options = {}) {
    const doAutofilter = options.doAutofilter !== false;
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
    const layout = this._resolveLayout(
      options,
      headers !== null,
      columnCount,
      headers,
      totalRows
    );
    sheet.layout = layout;
    const totalsOverrides = tableTotalsOverrides(layout.tables);
    if (totalsOverrides.length > 0) {
      const headerOffset = headers ? 1 : 0;
      const byRow = /* @__PURE__ */ new Map();
      for (const override of totalsOverrides) {
        const index = override.row - headerOffset;
        if (index < 0 || index >= rows.length) {
          throw new Error(
            `Table totals row ${override.row + 1} is outside the written rows (the sheet has ${totalRows} rows)`
          );
        }
        const list = byRow.get(index) ?? [];
        list.push(override);
        byRow.set(index, list);
      }
      for (const [index, list] of byRow) {
        rows[index] = withTableTotals(rows[index], list);
      }
    }
    const colWidths = this._resolveColumnWidths(
      columnCount,
      headers,
      rows,
      options,
      doAutofilter && headers !== null,
      layout.headerStyle ?? void 0
    );
    const colLetters = new Array(columnCount);
    for (let i = 0; i < columnCount; i++) colLetters[i] = columnIndexToLetter(i);
    const dimensionRef = totalRows > 0 && columnCount > 0 ? `A1:${colLetters[columnCount - 1]}${totalRows}` : "A1";
    writeXlsxSheetHead(buffer, {
      dimensionRef,
      columnCount,
      isFirstSheet: this._sheetCount === 1,
      hasHeaders: headers !== null,
      colWidths,
      freezePanes: layout.freezePanes,
      showGridLines: layout.showGridLines,
      showRowColHeaders: layout.showRowColHeaders,
      showZeroValues: layout.showZeroValues,
      rightToLeft: layout.rightToLeft,
      zoom: layout.zoom,
      tabColor: layout.tabColor,
      hiddenColumns: layout.hiddenColumns,
      columnOutline: layout.columnOutline
    });
    let rowNum = 0;
    if (headers) {
      rowNum++;
      buffer.writeString(`<row r="${rowNum}"${xlsxRowAttributes(0, layout)}>`);
      const headerStyle = this._headerStyleIndex(layout);
      const baseHeader = layout.headerStyle ?? { font: { bold: true } };
      for (let c = 0; c < headers.length; c++) {
        const rangeStyle = layout.styleRanges.styleFor(0, c);
        const style = rangeStyle ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle)) : headerStyle;
        this.writeStringCell(buffer, headers[c], "", style);
      }
      buffer.writeString("</row>");
    }
    for (const row of rows) {
      rowNum++;
      const sheetRow = rowNum - 1;
      buffer.writeString(`<row r="${rowNum}"${xlsxRowAttributes(sheetRow, layout)}>`);
      let prevCol = -1;
      for (let c = 0; c < row.length; c++) {
        const rangeStyle = layout.styleRanges.styleFor(sheetRow, c);
        prevCol = writeXlsxCell(
          buffer,
          this,
          row[c],
          c,
          prevCol,
          colLetters[c],
          rowNum,
          1,
          2,
          rangeStyle
        );
      }
      buffer.writeString("</row>");
    }
    this._applySheetFeatures(sheet, options, layout);
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
      registerDxf: (style) => this._styles.registerDxf(style),
      sheetProtection: layout.sheetProtection,
      autoFilterRef: layout.autoFilterRef,
      tableParts: sheet.tableParts
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
    const tableParts = this._sheets.flatMap(
      (sheet) => (sheet.tableParts ?? []).map((part) => ({ path: part.path }))
    );
    this._zip.addFile("[Content_Types].xml", buildXlsxContentTypes(this._sheets, commentParts, tableParts));
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
function writeXlsbRowHeader(sink, rowNum, startCol, endCol, options = {}) {
  const {
    heightTwips = 300,
    hidden = false,
    outlineLevel = 0,
    collapsed = false
  } = options;
  if (outlineLevel < 0 || outlineLevel > 7) {
    throw new RangeError(`XLSB outline level must be 0..7, got ${outlineLevel}`);
  }
  const flags = (heightTwips !== 300 ? 32 : 0) | (hidden ? 16 : 0) | (collapsed ? 8 : 0) | outlineLevel;
  sink.ensureCapacity(27);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(25);
  sink.writeUnsafeInt32LE(rowNum);
  sink.writeUnsafeInt32LE(0);
  sink.writeUnsafeByte(heightTwips & 255);
  sink.writeUnsafeByte(heightTwips >> 8 & 255);
  sink.writeUnsafeByte(0);
  sink.writeUnsafeByte(flags);
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
function encodeSstItem(value) {
  const text = typeof value === "string" ? value : value.text;
  const runs = typeof value === "string" ? [] : value.runs;
  const encoded = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    encoded[i * 2] = code & 255;
    encoded[i * 2 + 1] = code >> 8 & 255;
  }
  const payloadLength = 1 + 4 + encoded.length + (runs.length > 0 ? 4 + runs.length * 4 : 0);
  const payload = new Uint8Array(payloadLength);
  const view = new DataView(payload.buffer);
  payload[0] = runs.length > 0 ? 1 : 0;
  view.setUint32(1, text.length, true);
  payload.set(encoded, 5);
  if (runs.length > 0) {
    let offset = 5 + encoded.length;
    view.setUint32(offset, runs.length, true);
    offset += 4;
    for (const run of runs) {
      view.setUint16(offset, run.ich, true);
      view.setUint16(offset + 2, run.ifnt, true);
      offset += 4;
    }
  }
  return buildBrtRecord(19, payload);
}
function writeXlsbSharedStrings(sink, entries, countAll, countUnique) {
  sink.writeByte(159);
  sink.writeByte(1);
  sink.writeByte(8);
  sink.writeInt32LE(countAll);
  sink.writeInt32LE(countUnique);
  for (const value of entries) {
    sink.write(encodeSstItem(value));
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
var BRT_BEGIN_DXFS = 505;
var BRT_END_DXFS = 506;
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
  if ("auto" in color) return [1, 0, 0, 0, 0, 0, 0, 255];
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
function doubleBytes(value) {
  const buffer = new ArrayBuffer(8);
  new DataView(buffer).setFloat64(0, value, true);
  return [...new Uint8Array(buffer)];
}
function gradientFillPayload(fill) {
  const gradient = fill.gradient;
  const linear = (gradient.type ?? "linear") === "linear";
  const payload = [
    ...u32(40),
    ...colorBytes(void 0),
    ...colorBytes(void 0),
    ...u32(linear ? 0 : 1)
  ];
  const angle = linear && gradient.degree !== void 0 ? (360 - gradient.degree) % 360 : 0;
  payload.push(...doubleBytes(angle));
  if (linear) {
    payload.push(...doubleBytes(0), ...doubleBytes(0), ...doubleBytes(0), ...doubleBytes(0));
  } else {
    payload.push(
      ...doubleBytes(gradient.left ?? 0),
      ...doubleBytes(gradient.right ?? 1),
      ...doubleBytes(gradient.top ?? 0),
      ...doubleBytes(gradient.bottom ?? 1)
    );
  }
  payload.push(...u32(gradient.stops.length));
  for (const stop of gradient.stops) {
    payload.push(...colorBytes(stop.color), ...doubleBytes(stop.position));
  }
  return Uint8Array.from(payload);
}
function buildXlsbFillRecord(fill) {
  if (fill.gradient) {
    return buildBrtRecord(BRT_FILL, gradientFillPayload(fill));
  }
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
  for (const side of ["start", "end", "vertical", "horizontal"]) {
    if (border[side] !== void 0) {
      throw new Error(`XLSB has no record slot for the '${side}' border side; use XlsxWriter`);
    }
  }
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
  if (style?.quotePrefix) {
    throw new Error("XLSB has no record slot for quotePrefix; use XlsxWriter");
  }
  if (style?.pivotButton) {
    throw new Error("XLSB has no record slot for pivotButton; use XlsxWriter");
  }
  const alignment = style?.alignment;
  const protection = style?.protection;
  let rotation = 0;
  let indent = 0;
  let alignmentByte = 16;
  let flagsByte = 16;
  if (alignment) {
    if (alignment.relativeIndent !== void 0) {
      throw new Error("XLSB has no record slot for relativeIndent; use XlsxWriter");
    }
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
  /** Number-format id for a differential style (no new XF is allocated). */
  registerNumberFormat(format) {
    return this.numFmtIdFor(format);
  }
  /** Font-table id for one rich-text run (shared with cell-style fonts). */
  registerRichFont(font) {
    if (!font || Object.keys(font).length === 0) return 0;
    const key = JSON.stringify(font);
    return this.componentId(
      this.fontIds,
      key,
      buildXlsbFontRecord(font),
      this.fonts,
      font.bold && Object.keys(font).length === 1 ? 1 : null,
      () => this.nextFontId++
    );
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
  /**
   * Rewrite `styles.bin`, appending registered components to the template.
   * `dxfRecords` are the BrtBeginDXFs…BrtEndDXFs bytes for conditional
   * formats; they replace the template's empty DXF section.
   */
  toStylesBin(template, dxfRecords) {
    if (this.xfs.length === 0 && this.fonts.length === 0 && this.fills.length === 0 && this.borders.length === 0 && this.formatRecords.length === 0 && (!dxfRecords || dxfRecords.length === 0)) {
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
    if (dxfRecords && dxfRecords.length > 0) {
      const dxfs = section(BRT_BEGIN_DXFS, BRT_END_DXFS);
      parts.push(template.subarray(cellXfs.end.end, dxfs.begin.start));
      parts.push(dxfRecords);
      parts.push(template.subarray(dxfs.end.end));
    } else {
      parts.push(template.subarray(cellXfs.end.end));
    }
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
  TIME: { id: 66, min: 3, max: 3, fixed: 3, classes: "VVV" },
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
  SUBTOTAL: { id: 344, min: 2, max: 255, fixed: null, classes: "VR" },
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
function columnIndex2(letters) {
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
    const col = columnIndex2(match2[2]);
    const row = parseInt(match2[4], 10) - 1;
    if (col > 16383 || row > 1048575) {
      throw new RangeError(`Cell reference '${text}' is outside the Excel grid`);
    }
    return { row, col, colRelative: match2[1] === "", rowRelative: match2[3] === "" };
  }
  if (kind === "col") {
    const match2 = /^(\$?)([A-Za-z]{1,3})$/.exec(text);
    if (!match2) throw new Error(`Invalid XLSB column reference '${text}'`);
    return { row: 0, col: columnIndex2(match2[2]), rowRelative: false, colRelative: match2[1] === "" };
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
var DEFAULT_CONTEXT = { originCol: 0, originRow: 0, relativeOffsets: false };
function applyOrigin(part, context) {
  if (!context.relativeOffsets) return part;
  return {
    ...part,
    row: part.rowRelative ? part.row - context.originRow & 1048575 : part.row,
    col: part.colRelative ? part.col - context.originCol & 16383 : part.col
  };
}
function emitReference(text, operand, context = DEFAULT_CONTEXT) {
  const parts = text.split(":");
  const base = operand === "R" ? 32 : operand === "A" ? 96 : 64;
  if (parts.length === 1) {
    const part = applyOrigin(parseReferencePart(parts[0], "cell"), context);
    const relative2 = context.relativeOffsets && (part.colRelative || part.rowRelative);
    return [base | (relative2 ? 12 : 4), ...u322(part.row), ...u162(columnField(part))];
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
  const relative = context.relativeOffsets && (first.colRelative || first.rowRelative || last.colRelative || last.rowRelative);
  first = applyOrigin(first, context);
  last = applyOrigin(last, context);
  return [
    base | (relative ? 13 : 5),
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
function doubleBytes2(value) {
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
function emitNode(node, operand = "V", context = DEFAULT_CONTEXT) {
  switch (node.kind) {
    case "number": {
      if (Number.isInteger(node.value) && node.value >= 0 && node.value <= 65535) {
        return [30, ...u162(node.value)];
      }
      return [31, ...doubleBytes2(node.value)];
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
      return emitReference(node.value, operand, context);
    case "paren":
      return [...emitNode(node.value, operand, context), 21];
    case "percent":
      return [...emitNode(node.value, "V", context), 20];
    case "unary":
      return [...emitNode(node.value, "V", context), node.op === "+" ? 18 : 19];
    case "binary": {
      const argType = node.op === "," || node.op === " " || node.op === ":" ? "R" : "V";
      return [
        ...emitNode(node.left, argType, context),
        ...emitNode(node.right, argType, context),
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
      const encoded = node.args.map((arg, index) => emitNode(arg, functionArgumentClass(info, index), context));
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
function compileXlsbFormulaStats(formula, options = {}) {
  const text = formula.trim().replace(/^=/, "");
  if (!text) throw new Error("FormulaCell.formula must not be empty");
  if (text.length > 8192) throw new RangeError("XLSB formula text is limited to 8192 characters");
  const rawTokens = tokenize(text);
  let references = 0;
  for (let index = 0; index < rawTokens.length; index++) {
    if (ADDRESS_RE.test(rawTokens[index]) && rawTokens[index + 1] !== "(") references++;
  }
  const tree = new Parser(rawTokens).parse();
  const context = {
    originCol: options.origin?.col ?? 0,
    originRow: options.origin?.row ?? 0,
    relativeOffsets: options.relativeOffsets === true
  };
  let tokens = emitNode(tree, "V", context);
  if (containsVolatile(tree)) {
    tokens = [25, 1, 0, 0, ...tokens];
  }
  if (tokens.length === 0 || tokens.length > 16384) {
    throw new RangeError("XLSB formula token length exceeds the BIFF12 limit");
  }
  return { tokens: Uint8Array.from(tokens), references };
}
function compileXlsbFormula(formula, options = {}) {
  return compileXlsbFormulaStats(formula, options).tokens;
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
    cache = doubleBytes2(value);
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

// src/writers/core/xlsbLayout.ts
var BRT_PANE = 151;
var BRT_SEL = 152;
var BRT_BEGIN_COL_INFOS = 390;
var BRT_COL_INFO = 60;
var BRT_END_COL_INFOS = 391;
var PANE_ACTIVE = {
  bottomRight: 0,
  topRight: 1,
  bottomLeft: 2,
  topLeft: 3
};
function xlsbColorPayload(color) {
  if (color === void 0) return new Uint8Array(8);
  const key = colorKey(color);
  if (key === "auto") return Uint8Array.from([1, 64, 0, 0, 0, 0, 0, 0]);
  if (key.startsWith("rgb:")) {
    const value = key.slice(4);
    const alpha = parseInt(value.slice(0, 2), 16);
    const red = parseInt(value.slice(2, 4), 16);
    const green = parseInt(value.slice(4, 6), 16);
    const blue = parseInt(value.slice(6, 8), 16);
    return Uint8Array.from([5, 0, 0, 0, red, green, blue, alpha]);
  }
  if (key.startsWith("indexed:")) {
    const index = parseInt(key.slice(8), 10) & 255;
    return Uint8Array.from([3, index, 0, 0, 0, 0, 0, 255]);
  }
  const [, themeText, tintText] = key.split(":");
  const theme = parseInt(themeText, 10) & 255;
  const tint = Math.max(-32768, Math.min(32767, Math.round(Number(tintText) * 32767)));
  return Uint8Array.from([
    7,
    theme,
    tint & 255,
    tint >> 8 & 255,
    0,
    0,
    0,
    255
  ]);
}
function buildXlsbPaneRecord(freeze) {
  const xSplit = freeze?.xSplit ?? 0;
  const ySplit = freeze?.ySplit ?? 0;
  const active = PANE_ACTIVE[freeze?.activePane ?? "topLeft"];
  const flags = freeze ? 3 : 0;
  const payload = new Uint8Array(8 + 8 + 4 + 4 + 4 + 1);
  const view = new DataView(payload.buffer);
  view.setFloat64(0, xSplit, true);
  view.setFloat64(8, ySplit, true);
  view.setUint32(16, ySplit, true);
  view.setUint32(20, xSplit, true);
  view.setUint32(24, active, true);
  payload[28] = flags;
  return buildBrtRecord(BRT_PANE, payload);
}
function buildXlsbSelectionRecord(freeze) {
  const active = PANE_ACTIVE[freeze?.activePane ?? "topLeft"];
  const row = freeze?.ySplit ?? 0;
  const col = freeze?.xSplit ?? 0;
  const payload = new Uint8Array(20 + 16);
  const view = new DataView(payload.buffer);
  view.setUint32(0, active, true);
  view.setUint32(4, row, true);
  view.setUint32(8, col, true);
  view.setUint32(12, 0, true);
  view.setUint32(16, 1, true);
  view.setUint32(20, row, true);
  view.setUint32(24, row, true);
  view.setUint32(28, col, true);
  view.setUint32(32, col, true);
  return buildBrtRecord(BRT_SEL, payload);
}
var BRT_SHEET_PROTECTION = 535;
var BRT_LEGACY_DRAWING = 551;
var XLSB_PROTECTION_TEMPLATE_OFFSET = 221;
function buildXlsbSheetProtectionRecord(spec) {
  const allowed = (blocked) => blocked === false ? 1 : 0;
  const selection = (allow) => allow === false ? 0 : 1;
  const flags = [
    1,
    spec.objects === false ? 1 : 0,
    spec.scenarios === false ? 1 : 0,
    allowed(spec.formatCells),
    allowed(spec.formatColumns),
    allowed(spec.formatRows),
    allowed(spec.insertColumns),
    allowed(spec.insertRows),
    allowed(spec.insertHyperlinks),
    allowed(spec.deleteColumns),
    allowed(spec.deleteRows),
    selection(spec.selectLockedCells),
    allowed(spec.sort),
    allowed(spec.autoFilter),
    allowed(spec.pivotTables),
    selection(spec.selectUnlockedCells)
  ];
  const payload = new Uint8Array(2 + flags.length * 4);
  const view = new DataView(payload.buffer);
  view.setUint16(0, spec.password ? passwordHash(spec.password) : 0, true);
  flags.forEach((value, index) => view.setUint32(2 + index * 4, value, true));
  return buildBrtRecord(BRT_SHEET_PROTECTION, payload);
}
function applyXlsbSheetProtection(template, spec) {
  if (!spec || spec.enabled === false) return;
  const record = buildXlsbSheetProtectionRecord(spec);
  if (XLSB_PROTECTION_TEMPLATE_OFFSET + record.length > template.length) {
    throw new Error("XLSB sheet template is too small for the protection record");
  }
  template.set(record, XLSB_PROTECTION_TEMPLATE_OFFSET);
}
function buildXlsbLegacyDrawingRecord(relId) {
  const payload = new Uint8Array(4 + relId.length * 2);
  const view = new DataView(payload.buffer);
  view.setUint32(0, relId.length, true);
  for (let index = 0; index < relId.length; index++) {
    view.setUint16(4 + index * 2, relId.charCodeAt(index), true);
  }
  return buildBrtRecord(BRT_LEGACY_DRAWING, payload);
}
function buildXlsbColumnRecords(layout) {
  const hidden = new Set(layout.hiddenColumns);
  let maxCol = Math.max(layout.endCol - 1, layout.startCol);
  for (const col of hidden) maxCol = Math.max(maxCol, col);
  if (layout.outline) for (const col of layout.outline.indices) maxCol = Math.max(maxCol, col);
  const parts = [buildBrtRecord(BRT_BEGIN_COL_INFOS, new Uint8Array(0))];
  for (let col = layout.startCol; col <= maxCol; col++) {
    const width = layout.widths[col] ?? 8.43;
    let flags = 2;
    if (hidden.has(col)) flags |= 1;
    if (layout.outline && col < layout.outline.level.length) {
      flags |= layout.outline.level[col] << 8;
      if (layout.outline.collapsed[col]) flags |= 4096;
    }
    const payload = new Uint8Array(18);
    const view = new DataView(payload.buffer);
    view.setUint32(0, col, true);
    view.setUint32(4, col, true);
    view.setUint32(8, Math.max(0, Math.min(255 * 256, Math.round(width * 256))), true);
    view.setUint32(12, 0, true);
    view.setUint16(16, flags, true);
    parts.push(buildBrtRecord(BRT_COL_INFO, payload));
  }
  parts.push(buildBrtRecord(BRT_END_COL_INFOS, new Uint8Array(0)));
  return concat(parts);
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
  greaterThan: 4,
  lessThan: 5,
  greaterThanOrEqual: 6,
  lessThanOrEqual: 7
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
function columnIndex3(letters) {
  let value = 0;
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value - 1;
}
function parseCell(text) {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(text.trim());
  if (!match) throw new Error(`Invalid cell reference '${text}'`);
  const col = columnIndex3(match[1]);
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
function validationFormula(text, origin) {
  if (text === void 0 || text === "") return new Uint8Array(8);
  const tokens = compileXlsbFormula(text, { origin, relativeOffsets: true });
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
    const origin = { col: ranges[0].c1, row: ranges[0].r1 };
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
    payload.push(...validationFormula(formula1, origin));
    payload.push(...validationFormula(rule.formula2, origin));
    parts.push(buildBrtRecord(BRT_DVAL, Uint8Array.from(payload)));
  }
  parts.push(buildBrtRecord(BRT_END_DVALS, new Uint8Array(0)));
  return concat(parts);
}

// src/writers/core/xlsbConditionalFormat.ts
var BRT_BEGIN_DXFS2 = 505;
var BRT_END_DXFS2 = 506;
var BRT_DXF = 507;
var BRT_BEGIN_CF = 461;
var BRT_END_CF = 462;
var BRT_BEGIN_CF_RULE = 463;
var BRT_END_CF_RULE = 464;
var BRT_BEGIN_ICON_SET = 465;
var BRT_END_ICON_SET = 466;
var BRT_BEGIN_DATA_BAR = 467;
var BRT_END_DATA_BAR = 468;
var BRT_BEGIN_COLOR_SCALE = 469;
var BRT_END_COLOR_SCALE = 470;
var BRT_CFVO = 471;
var BRT_COLOR = 564;
var CF_CELL_IS = 1;
var CF_EXPRESSION = 2;
var CF_TOP10 = 5;
var CF_COLOR_SCALE = 3;
var CF_DATA_BAR = 4;
var CF_ICON_SET = 6;
var TEMPLATE_FORMULA = 1;
var TEMPLATE_FILTER = 5;
var TEMPLATE_UNIQUE = 7;
var TEMPLATE_DUPLICATE = 27;
var CELL_IS_CODES = {
  between: 1,
  notBetween: 2,
  equal: 3,
  notEqual: 4,
  greaterThan: 5,
  lessThan: 6,
  greaterThanOrEqual: 7,
  lessThanOrEqual: 8
};
var TEXT_OPERATORS = {
  containsText: 0,
  notContainsText: 1,
  beginsWith: 2,
  endsWith: 3
};
var CFVO_CODES = {
  num: 1,
  min: 2,
  max: 3,
  percent: 4,
  percentile: 5,
  formula: 7
};
var ICON_SET_CODES = {
  "3Arrows": 0,
  "3ArrowsGray": 1,
  "3Flags": 2,
  "3TrafficLights1": 3,
  "3TrafficLights2": 4,
  "3Signs": 5,
  "3Symbols": 6,
  "3Symbols2": 7,
  "4Arrows": 8,
  "4ArrowsGray": 9,
  "4RedToBlack": 10,
  "4Rating": 11,
  "4TrafficLights": 12,
  "5Arrows": 13,
  "5ArrowsGray": 14,
  "5Rating": 15,
  "5Quarters": 16
};
var PROP_FILL_PATTERN = 0;
var PROP_FOREGROUND = 1;
var PROP_BACKGROUND = 2;
var PROP_FONT_COLOR = 5;
var PROP_BORDER_TOP = 6;
var PROP_BORDER_BOTTOM = 7;
var PROP_BORDER_LEFT = 8;
var PROP_BORDER_RIGHT = 9;
var PROP_HORIZONTAL = 15;
var PROP_WRAP = 20;
var PROP_FONT_NAME = 24;
var PROP_BOLD = 25;
var PROP_UNDERLINE = 26;
var PROP_ITALIC = 28;
var PROP_STRIKE = 29;
var PROP_SIZE = 36;
var PROP_NUMBER_FORMAT = 38;
var PROP_NUMBER_FORMAT_ID = 41;
var PROP_LOCKED = 43;
var PROP_HIDDEN = 44;
var DXF_FLAGS = 32768;
var STOP_IF_TRUE = 2;
var PERCENT_FLAG = 8;
var BOTTOM_FLAG = 16;
var PATTERN_CODES = {
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
var BORDER_CODES = {
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
var UNDERLINE_CODES = {
  single: 1,
  double: 2,
  singleAccounting: 33,
  doubleAccounting: 34
};
var HORIZONTAL_CODES = {
  general: 0,
  left: 1,
  center: 2,
  right: 3,
  fill: 4,
  justify: 5,
  centerContinuous: 6,
  distributed: 7
};
function u164(value) {
  const v = value & 65535;
  return [v & 255, v >>> 8 & 255];
}
function u324(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function i32(value) {
  return u324(value >>> 0);
}
function doubleBytes3(value) {
  const buffer = new ArrayBuffer(8);
  new DataView(buffer).setFloat64(0, value, true);
  return [...new Uint8Array(buffer)];
}
function wideString3(text) {
  const out = [...u324(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function nullableString2(text) {
  return text === void 0 ? [255, 255, 255, 255] : wideString3(text);
}
function cfColor(color, opaqueRgb = false) {
  if (color === void 0) return [0, 0, 0, 0, 0, 0, 0, 0];
  const key = colorKey(color);
  if (key === "auto") return [1, 64, 0, 0, 0, 0, 0, 0];
  if (key.startsWith("rgb:")) {
    const value = key.slice(4);
    const alpha = opaqueRgb ? 255 : parseInt(value.slice(0, 2), 16);
    return [
      5,
      255,
      0,
      0,
      parseInt(value.slice(2, 4), 16),
      parseInt(value.slice(4, 6), 16),
      parseInt(value.slice(6, 8), 16),
      alpha
    ];
  }
  if (key.startsWith("indexed:")) {
    return [3, parseInt(key.slice(8), 10) & 255, 0, 0, 0, 0, 0, 255];
  }
  const [, themeText, tintText] = key.split(":");
  const tint = Math.max(-32768, Math.min(32767, Math.round(Number(tintText) * 32767)));
  return [
    7,
    parseInt(themeText, 10) & 255,
    tint & 255,
    tint >> 8 & 255,
    0,
    0,
    0,
    255
  ];
}
function prop(kind, data) {
  return {
    kind,
    bytes: [...u164(kind), ...u164(4 + data.length), ...data]
  };
}
function unsupported(name) {
  throw new Error(
    `${name} is not supported in XLSB conditional formats; use XlsxWriter for it`
  );
}
function fontProps(font) {
  if (font.vertAlign !== void 0) unsupported("Font.vertAlign");
  if (font.family !== void 0) unsupported("Font.family");
  if (font.scheme !== void 0) unsupported("Font.scheme");
  if (font.outline !== void 0) unsupported("Font.outline");
  if (font.shadow !== void 0) unsupported("Font.shadow");
  if (font.condense !== void 0) unsupported("Font.condense");
  if (font.extend !== void 0) unsupported("Font.extend");
  if (font.charset !== void 0) unsupported("Font.charset");
  const props = [];
  if (font.color !== void 0) props.push(prop(PROP_FONT_COLOR, cfColor(font.color)));
  if (font.name !== void 0) props.push(prop(PROP_FONT_NAME, wideString3(font.name)));
  if (font.bold !== void 0) props.push(prop(PROP_BOLD, u164(font.bold ? 700 : 400)));
  if (font.underline !== void 0) {
    const key = font.underline === true ? "single" : font.underline;
    props.push(prop(PROP_UNDERLINE, u164(key === false ? 0 : UNDERLINE_CODES[key] ?? 0)));
  }
  if (font.italic !== void 0) props.push(prop(PROP_ITALIC, [font.italic ? 1 : 0]));
  if (font.strike !== void 0) props.push(prop(PROP_STRIKE, [font.strike ? 1 : 0]));
  if (font.size !== void 0) props.push(prop(PROP_SIZE, u324(Math.round(font.size * 20))));
  return props;
}
function fillProps(fill) {
  if (fill.gradient) unsupported("GradientFill");
  const props = [];
  if (fill.patternType !== void 0) {
    props.push(prop(PROP_FILL_PATTERN, [PATTERN_CODES[fill.patternType] ?? 0]));
  }
  if (fill.fgColor !== void 0) props.push(prop(PROP_FOREGROUND, cfColor(fill.fgColor)));
  if (fill.bgColor !== void 0) {
    props.push(prop(PROP_BACKGROUND, cfColor(fill.bgColor)));
  } else if (fill.patternType === "solid" && fill.fgColor !== void 0) {
    props.push(prop(PROP_BACKGROUND, cfColor(fill.fgColor)));
  }
  return props;
}
function borderProps(border) {
  for (const side of ["diagonal", "start", "end", "vertical", "horizontal"]) {
    if (border[side] !== void 0) unsupported(`Border.${side}`);
  }
  const props = [];
  const sides = [
    ["top", PROP_BORDER_TOP],
    ["bottom", PROP_BORDER_BOTTOM],
    ["left", PROP_BORDER_LEFT],
    ["right", PROP_BORDER_RIGHT]
  ];
  for (const [name, kind] of sides) {
    const side = border[name];
    const style = side?.style ?? "none";
    const color = side?.color === void 0 ? [1, 64, 0, 0, 0, 0, 0, 0] : cfColor(side.color);
    props.push(prop(kind, [...color, ...u164(BORDER_CODES[style] ?? 0)]));
  }
  return props;
}
function alignmentProps(alignment) {
  for (const name of [
    "vertical",
    "shrinkToFit",
    "indent",
    "textRotation",
    "relativeIndent",
    "justifyLastLine",
    "readingOrder"
  ]) {
    if (alignment[name] !== void 0) unsupported(`Alignment.${name}`);
  }
  const props = [];
  if (alignment.horizontal !== void 0) {
    props.push(prop(PROP_HORIZONTAL, [HORIZONTAL_CODES[alignment.horizontal] ?? 0]));
  }
  if (alignment.wrapText !== void 0) {
    props.push(prop(PROP_WRAP, [alignment.wrapText ? 1 : 0]));
  }
  return props;
}
function buildXlsbDxfPayload(style, registerNumFmt) {
  const props = [];
  if (style.font) props.push(...fontProps(style.font));
  if (style.fill) props.push(...fillProps(style.fill));
  if (style.border) props.push(...borderProps(style.border));
  if (style.alignment) props.push(...alignmentProps(style.alignment));
  if (style.protection?.locked !== void 0) {
    props.push(prop(PROP_LOCKED, [style.protection.locked ? 1 : 0]));
  }
  if (style.protection?.hidden !== void 0) {
    props.push(prop(PROP_HIDDEN, [style.protection.hidden ? 1 : 0]));
  }
  if (style.numFmt !== void 0) {
    const id = registerNumFmt(style.numFmt);
    props.push({
      kind: PROP_NUMBER_FORMAT - 0.5,
      bytes: [...u164(PROP_NUMBER_FORMAT_ID), ...u164(6), ...u164(id & 65535)]
    });
    props.push(prop(PROP_NUMBER_FORMAT, wideString3(style.numFmt)));
  }
  props.sort((left, right) => left.kind - right.kind);
  return Uint8Array.from([
    ...u164(DXF_FLAGS),
    ...u164(0),
    ...u164(props.length),
    ...props.flatMap((item) => item.bytes)
  ]);
}
var XlsbDxfRegistry = class {
  constructor(registerNumFmt) {
    this.registerNumFmt = registerNumFmt;
    this.ids = /* @__PURE__ */ new Map();
    this.payloads = [];
  }
  register(style) {
    if (!style) return -1;
    const key = JSON.stringify(style);
    const existing = this.ids.get(key);
    if (existing !== void 0) return existing;
    const id = this.payloads.length;
    this.ids.set(key, id);
    this.payloads.push(buildXlsbDxfPayload(style, this.registerNumFmt));
    return id;
  }
  get size() {
    return this.payloads.length;
  }
  /** BrtBeginDXFs / BrtDXF* / BrtEndDXFs, for insertion in styles.bin. */
  toRecords() {
    if (this.payloads.length === 0) return new Uint8Array(0);
    const parts = [
      buildBrtRecord(BRT_BEGIN_DXFS2, Uint8Array.from(u324(this.payloads.length)))
    ];
    for (const payload of this.payloads) {
      parts.push(buildBrtRecord(BRT_DXF, payload));
    }
    parts.push(buildBrtRecord(BRT_END_DXFS2, new Uint8Array(0)));
    return concat(parts);
  }
};
function cfvoRecord(type, value, gte, origin) {
  const code = CFVO_CODES[type];
  if (type === "formula") {
    if (value === void 0) throw new Error("A formula CFVO needs a value");
    const compiled = compileXlsbFormulaStats(value, { origin, relativeOffsets: true });
    const size = compiled.tokens.length + 2 * compiled.references;
    return buildBrtRecord(BRT_CFVO, Uint8Array.from([
      ...u324(code),
      ...doubleBytes3(0),
      ...u324(gte ? 1 : 0),
      ...u324(gte ? 1 : 0),
      ...u324(size),
      ...u324(compiled.tokens.length),
      ...compiled.tokens,
      ...u324(0)
    ]));
  }
  const number = type === "min" || type === "max" ? 0 : Number(value ?? 0);
  if (!Number.isFinite(number)) throw new RangeError("CFVO values must be finite");
  return buildBrtRecord(BRT_CFVO, Uint8Array.from([
    ...u324(code),
    ...doubleBytes3(number),
    ...u324(gte ? 1 : 0),
    ...u324(gte ? 1 : 0),
    ...u324(0)
  ]));
}
function colorRecord(color) {
  return buildBrtRecord(BRT_COLOR, Uint8Array.from(cfColor(color, true)));
}
function rulePayload(kind, template, dxfId, priority, param, flags, formulas, text, origin) {
  const compiled = formulas.map((formula) => compileXlsbFormulaStats(formula, { origin, relativeOffsets: true }));
  const sizes = compiled.map((item) => item.tokens.length + 2 * item.references);
  while (sizes.length < 3) sizes.push(0);
  const payload = [
    ...u324(kind),
    ...u324(template),
    ...i32(dxfId),
    ...u324(priority),
    ...u324(param),
    ...u324(0),
    ...u324(0),
    ...u164(flags),
    ...u324(sizes[0]),
    ...u324(sizes[1]),
    ...u324(sizes[2]),
    ...nullableString2(text)
  ];
  for (const item of compiled) {
    payload.push(...u324(item.tokens.length), ...item.tokens, ...u324(0));
  }
  return Uint8Array.from(payload);
}
function scaleRule(rule, priority, origin) {
  const kind = rule.type === "colorScale" ? CF_COLOR_SCALE : rule.type === "dataBar" ? CF_DATA_BAR : CF_ICON_SET;
  const template = kind === CF_COLOR_SCALE ? 2 : kind === CF_DATA_BAR ? 3 : 4;
  const head = buildBrtRecord(BRT_BEGIN_CF_RULE, Uint8Array.from([
    ...u324(kind),
    ...u324(template),
    ...i32(-1),
    ...u324(priority),
    ...u324(0),
    ...u324(0),
    ...u324(0),
    ...u164(0),
    ...u324(0),
    ...u324(0),
    ...u324(0),
    ...i32(-1)
  ]));
  if (rule.type === "colorScale") {
    if (rule.colors.length < 2) throw new Error("A color scale needs at least two colors");
    const parts2 = [buildBrtRecord(BRT_BEGIN_COLOR_SCALE, new Uint8Array(0))];
    const cfvo = rule.cfvo;
    if (cfvo.length === 2) {
      parts2.push(cfvoRecord(cfvo[0].type, cfvo[0].value, false, origin));
      parts2.push(cfvoRecord(cfvo[1].type, cfvo[1].value, false, origin));
    } else {
      const first = cfvo[0] ?? { type: "min" };
      const middle = cfvo[1] ?? { type: "percentile", value: "50" };
      const last = cfvo[2] ?? { type: "max" };
      parts2.push(cfvoRecord(first.type, first.value, false, origin));
      parts2.push(cfvoRecord(middle.type, middle.value, false, origin));
      parts2.push(cfvoRecord(last.type, last.value, false, origin));
    }
    for (const color of rule.colors) parts2.push(colorRecord(color));
    parts2.push(buildBrtRecord(BRT_END_COLOR_SCALE, new Uint8Array(0)));
    return concat([head, ...parts2, buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0))]);
  }
  if (rule.type === "dataBar") {
    const cfvo = rule.cfvo;
    const first = cfvo[0] ?? { type: "min" };
    const last = cfvo[1] ?? { type: "max" };
    return concat([
      head,
      buildBrtRecord(BRT_BEGIN_DATA_BAR, Uint8Array.from([
        10,
        90,
        rule.showValue === false ? 0 : 1
      ])),
      cfvoRecord(first.type, first.value, false, origin),
      cfvoRecord(last.type, last.value, false, origin),
      colorRecord(rule.color),
      buildBrtRecord(BRT_END_DATA_BAR, new Uint8Array(0)),
      buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0))
    ]);
  }
  const iconSet = rule.iconSet ?? "3TrafficLights1";
  const iconCode = ICON_SET_CODES[iconSet];
  if (iconCode === void 0) throw new Error(`Unknown XLSB icon set '${iconSet}'`);
  const count = iconSet.startsWith("5") ? 5 : iconSet.startsWith("4") ? 4 : 3;
  const defaults = Array.from({ length: count - 1 }, (_, index) => ({
    type: "percent",
    value: String(Math.round((index + 1) * 100 / count))
  }));
  const supplied = rule.cfvo ?? [];
  const thresholds = supplied.length >= count - 1 ? supplied.slice(0, count - 1) : [...supplied, ...defaults.slice(supplied.length)];
  const values = [
    { type: "percent", value: "0" },
    ...thresholds
  ];
  const flags = (rule.showValue === false ? 2 : 0) | (rule.reverse ? 4 : 0);
  const parts = [
    head,
    buildBrtRecord(BRT_BEGIN_ICON_SET, Uint8Array.from([...u324(iconCode), ...u164(flags)]))
  ];
  for (const item of values) parts.push(cfvoRecord(item.type, item.value, true, origin));
  parts.push(buildBrtRecord(BRT_END_ICON_SET, new Uint8Array(0)));
  parts.push(buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0)));
  return concat(parts);
}
function textFormula(rule, corner) {
  const quoted = `"${rule.text.replace(/"/g, '""')}"`;
  switch (rule.type) {
    case "containsText":
      return `NOT(ISERROR(SEARCH(${quoted},${corner})))`;
    case "notContainsText":
      return `ISERROR(SEARCH(${quoted},${corner}))`;
    case "beginsWith":
      return `LEFT(${corner},${rule.text.length})=${quoted}`;
    case "endsWith":
      return `RIGHT(${corner},${rule.text.length})=${quoted}`;
  }
}
function ruleBin(rule, sqref, priority, dxfRegistry) {
  const start = parseRangeRef(sqref.trim().split(/\s+/)[0]);
  const origin = { col: start.c1, row: start.r1 };
  if (rule.type === "colorScale" || rule.type === "dataBar" || rule.type === "iconSet") {
    return scaleRule(rule, priority, origin);
  }
  const corner = `${columnLetter(start.c1)}${start.r1 + 1}`;
  const dxfId = dxfRegistry.register(rule.style);
  const stop = "stopIfTrue" in rule && rule.stopIfTrue ? STOP_IF_TRUE : 0;
  let payload;
  if (rule.type === "cellIs") {
    const operator = CELL_IS_CODES[rule.operator];
    if (operator === void 0) {
      throw new Error(`CellIs operator '${rule.operator}' is not supported in XLSB`);
    }
    payload = rulePayload(
      CF_CELL_IS,
      0,
      dxfId,
      priority,
      operator,
      stop,
      rule.formula,
      void 0,
      origin
    );
  } else if (rule.type === "expression") {
    payload = rulePayload(
      CF_EXPRESSION,
      TEMPLATE_FORMULA,
      dxfId,
      priority,
      0,
      stop,
      [rule.formula],
      void 0,
      origin
    );
  } else if (rule.type === "top10") {
    payload = rulePayload(
      CF_TOP10,
      TEMPLATE_FILTER,
      dxfId,
      priority,
      rule.rank ?? 10,
      stop | (rule.percent ? PERCENT_FLAG : 0) | (rule.bottom ? BOTTOM_FLAG : 0),
      [],
      void 0,
      origin
    );
  } else if (rule.type === "duplicateValues" || rule.type === "uniqueValues") {
    payload = rulePayload(
      CF_EXPRESSION,
      rule.type === "uniqueValues" ? TEMPLATE_UNIQUE : TEMPLATE_DUPLICATE,
      dxfId,
      priority,
      0,
      stop,
      [],
      void 0,
      origin
    );
  } else if (rule.type === "containsText" || rule.type === "notContainsText" || rule.type === "beginsWith" || rule.type === "endsWith") {
    const template = TEXT_OPERATORS[rule.type];
    payload = rulePayload(
      CF_EXPRESSION,
      8,
      dxfId,
      priority,
      template,
      stop,
      [textFormula(rule, corner)],
      rule.text,
      origin
    );
  } else {
    throw new Error(`Unsupported XLSB conditional format rule '${rule.type}'`);
  }
  return concat([
    buildBrtRecord(BRT_BEGIN_CF_RULE, payload),
    buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0))
  ]);
}
function columnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}
function buildXlsbConditionalFormatRecords(formats, dxfRegistry, startPriority = 1) {
  const parts = [];
  let priority = startPriority;
  for (const format of formats) {
    if (!format.rules || format.rules.length === 0) continue;
    const ranges = format.sqref.trim().split(/\s+/).map(parseRangeRef);
    if (ranges.length === 0) throw new Error("Conditional formatting sqref must not be empty");
    const head = [
      ...u324(format.rules.length),
      ...u324(0),
      ...u324(ranges.length)
    ];
    for (const range of ranges) {
      head.push(...u324(range.r1), ...u324(range.r2), ...u324(range.c1), ...u324(range.c2));
    }
    parts.push(buildBrtRecord(BRT_BEGIN_CF, Uint8Array.from(head)));
    for (const rule of format.rules) {
      parts.push(ruleBin(rule, format.sqref, priority++, dxfRegistry));
    }
    parts.push(buildBrtRecord(BRT_END_CF, new Uint8Array(0)));
  }
  return concat(parts);
}

// src/writers/core/xlsbComments.ts
var BEGIN_COMMENTS = 628;
var END_COMMENTS = 629;
var BEGIN_AUTHORS = 630;
var END_AUTHORS = 631;
var AUTHOR = 632;
var BEGIN_LIST = 633;
var END_LIST = 634;
var BEGIN_COMMENT = 635;
var END_COMMENT = 636;
var TEXT = 637;
function u325(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function wideString4(text) {
  const out = [...u325(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function deterministicGuid(text) {
  let hash = 2166136261;
  const bytes = new Uint8Array(16);
  for (let block = 0; block < 4; block++) {
    for (let index = 0; index < text.length; index++) {
      hash ^= text.charCodeAt(index) + block * 31;
      hash = Math.imul(hash, 16777619) >>> 0;
    }
    bytes[block * 4] = hash & 255;
    bytes[block * 4 + 1] = hash >>> 8 & 255;
    bytes[block * 4 + 2] = hash >>> 16 & 255;
    bytes[block * 4 + 3] = hash >>> 24 & 255;
  }
  bytes[6] = bytes[6] & 15 | 80;
  bytes[8] = bytes[8] & 63 | 128;
  return bytes;
}
function commentGuid(partName, ref) {
  return deterministicGuid(`${partName}:${ref}`);
}
function buildXlsbCommentsPart(comments, partName) {
  if (comments.length === 0) return new Uint8Array(0);
  const authors = [];
  const authorIds = /* @__PURE__ */ new Map();
  for (const comment of comments) {
    const author = comment.author ?? "Author";
    if (!authorIds.has(author)) {
      authorIds.set(author, authors.length);
      authors.push(author);
    }
  }
  const parts = [
    buildBrtRecord(BEGIN_COMMENTS, new Uint8Array(0)),
    buildBrtRecord(BEGIN_AUTHORS, new Uint8Array(0))
  ];
  for (const author of authors) {
    parts.push(buildBrtRecord(AUTHOR, Uint8Array.from(wideString4(author))));
  }
  parts.push(buildBrtRecord(END_AUTHORS, new Uint8Array(0)));
  parts.push(buildBrtRecord(BEGIN_LIST, new Uint8Array(0)));
  for (const comment of comments) {
    const { row, col } = parseCellRef(comment.ref);
    const authorId = authorIds.get(comment.author ?? "Author");
    const guid = commentGuid(partName, comment.ref);
    const header = Uint8Array.from([
      ...u325(authorId),
      ...u325(row),
      ...u325(row),
      ...u325(col),
      ...u325(col),
      ...guid
    ]);
    parts.push(buildBrtRecord(BEGIN_COMMENT, header));
    parts.push(buildBrtRecord(TEXT, Uint8Array.from([0, ...wideString4(comment.text)])));
    parts.push(buildBrtRecord(END_COMMENT, new Uint8Array(0)));
  }
  parts.push(buildBrtRecord(END_LIST, new Uint8Array(0)));
  parts.push(buildBrtRecord(END_COMMENTS, new Uint8Array(0)));
  return concat(parts);
}
function buildXlsbCommentsVml(comments) {
  return buildXlsxCommentsVml(comments);
}

// src/writers/core/xlsbTable.ts
var BRT_AC_BEGIN = 37;
var BRT_FRT_BEGIN = 3072;
var BRT_AC_END = 38;
var BRT_BEGIN_LIST = 343;
var BRT_END_LIST = 344;
var BRT_BEGIN_LIST_COLS = 345;
var BRT_END_LIST_COLS = 346;
var BRT_BEGIN_LIST_COL = 347;
var BRT_END_LIST_COL = 348;
var BRT_LIST_TR_FMLA = 352;
var BRT_TABLE_STYLE_CLIENT = 513;
var BRT_BEGIN_AUTOFILTER = 161;
var BRT_END_AUTOFILTER = 162;
var BRT_BEGIN_LIST_PARTS = 660;
var BRT_LIST_PART = 661;
var BRT_END_LIST_PARTS = 662;
var TOTALS_CODES = {
  average: 1,
  count: 2,
  countNums: 3,
  max: 4,
  min: 5,
  sum: 6,
  stdDev: 7,
  var: 8,
  custom: 9
};
function u326(value) {
  const v = value >>> 0;
  return [v & 255, v >>> 8 & 255, v >>> 16 & 255, v >>> 24 & 255];
}
function i322(value) {
  return u326(value >>> 0);
}
function wideString5(text) {
  const out = [...u326(text.length)];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out.push(code & 255, code >>> 8 & 255);
  }
  return out;
}
function nullableString3(text) {
  return text === void 0 ? [255, 255, 255, 255] : wideString5(text);
}
function rangeBytes(range) {
  return [...u326(range.r1), ...u326(range.r2), ...u326(range.c1), ...u326(range.c2)];
}
function acuid(key) {
  return concat([
    buildBrtRecord(BRT_AC_BEGIN, Uint8Array.from([1, 0, 2, 16, 0, 128])),
    buildBrtRecord(BRT_FRT_BEGIN, deterministicGuid(key)),
    buildBrtRecord(BRT_AC_END, new Uint8Array(0))
  ]);
}
function buildXlsbTablePart(table, tableId) {
  const parts = [acuid(`${tableId}:${table.name}`)];
  const beginList = [
    ...rangeBytes(table.ref),
    ...u326(0),
    // lt = LTRANGE
    ...u326(tableId),
    ...u326(table.headerRow ? 1 : 0),
    ...u326(table.totalsRow ? 1 : 0),
    ...u326(table.totalsRow ? 1 : 0)
    // fShowTotals
  ];
  for (let index = 0; index < 6; index++) beginList.push(...i322(-1));
  beginList.push(...u326(0));
  beginList.push(...nullableString3(table.name));
  beginList.push(...nullableString3(table.name));
  beginList.push(...u326(0));
  for (let index = 0; index < 3; index++) beginList.push(...nullableString3(void 0));
  parts.push(buildBrtRecord(BRT_BEGIN_LIST, Uint8Array.from(beginList)));
  parts.push(acuid(`${tableId}:${table.name}:list`));
  if (table.headerRow) {
    const filterRef = tableAutoFilterRef(table);
    if (filterRef !== null) {
      const endRow = table.totalsRow ? table.ref.r2 - 1 : table.ref.r2;
      parts.push(buildBrtRecord(BRT_BEGIN_AUTOFILTER, Uint8Array.from([
        ...u326(table.ref.r1),
        ...u326(endRow),
        ...u326(table.ref.c1),
        ...u326(table.ref.c2)
      ])));
      parts.push(buildBrtRecord(BRT_END_AUTOFILTER, new Uint8Array(0)));
    }
  }
  parts.push(buildBrtRecord(BRT_BEGIN_LIST_COLS, Uint8Array.from(u326(table.columns.length))));
  table.columns.forEach((column, index) => {
    const number = index + 1;
    parts.push(acuid(`${tableId}:${table.name}:${number}`));
    const code = column.totalsRowFunction === void 0 ? 0 : TOTALS_CODES[column.totalsRowFunction] ?? 0;
    const columnPayload = [
      ...u326(number),
      ...u326(code),
      ...i322(-1),
      ...i322(-1),
      ...i322(-1),
      ...u326(0),
      // idqsif
      ...nullableString3(void 0),
      ...nullableString3(column.name),
      ...nullableString3(column.totalsRowLabel),
      ...nullableString3(void 0),
      ...nullableString3(void 0),
      ...nullableString3(void 0)
    ];
    parts.push(buildBrtRecord(BRT_BEGIN_LIST_COL, Uint8Array.from(columnPayload)));
    if (code === TOTALS_CODES.custom && column.totalsRowFormula !== void 0) {
      const compiled = compileXlsbFormulaStats(column.totalsRowFormula);
      parts.push(buildBrtRecord(BRT_LIST_TR_FMLA, Uint8Array.from([
        0,
        ...u326(compiled.tokens.length),
        ...compiled.tokens,
        ...u326(0)
      ])));
    }
    parts.push(buildBrtRecord(BRT_END_LIST_COL, new Uint8Array(0)));
  });
  parts.push(buildBrtRecord(BRT_END_LIST_COLS, new Uint8Array(0)));
  let flags = 0;
  if (table.showFirstColumn) flags |= 1;
  if (table.showLastColumn) flags |= 2;
  if (table.showRowStripes) flags |= 4;
  if (table.showColumnStripes) flags |= 8;
  parts.push(buildBrtRecord(BRT_TABLE_STYLE_CLIENT, Uint8Array.from([
    flags & 255,
    flags >>> 8 & 255,
    ...nullableString3(table.styleName)
  ])));
  parts.push(buildBrtRecord(BRT_END_LIST, new Uint8Array(0)));
  return concat(parts);
}
function buildXlsbListParts(relIds) {
  if (relIds.length === 0) return new Uint8Array(0);
  const parts = [
    buildBrtRecord(BRT_BEGIN_LIST_PARTS, Uint8Array.from(u326(relIds.length)))
  ];
  for (const relId of relIds) {
    parts.push(buildBrtRecord(BRT_LIST_PART, Uint8Array.from(wideString5(relId))));
  }
  parts.push(buildBrtRecord(BRT_END_LIST_PARTS, new Uint8Array(0)));
  return concat(parts);
}

// src/writers/core/theme.ts
var XLSX_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="\u7B49\u7EBF Light"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="\uB9D1\uC740 \uACE0\uB515"/><a:font script="Hans" typeface="\u7B49\u7EBF"/><a:font script="Hant" typeface="\u65B0\u7D30\u660E\u9AD4"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>

`;

// src/browser/BrowserXlsbWriter.ts
var INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;
function withTableTotals2(row, overrides) {
  let result = null;
  for (const override of overrides) {
    if (override.col < 0) continue;
    const original = row[override.col];
    const record = original !== null && typeof original === "object" && !(original instanceof Date) ? original : null;
    const style = record && record.style ? record.style : void 0;
    const cachedValue = record?.value;
    const cached = typeof cachedValue === "string" || typeof cachedValue === "number" || typeof cachedValue === "boolean" || cachedValue instanceof Date ? cachedValue : typeof original === "number" ? original : void 0;
    const replacement = override.text !== void 0 ? style ? { value: override.text, style } : override.text : {
      formula: override.formula,
      ...cached !== void 0 ? { value: cached } : {},
      ...style ? { style } : {}
    };
    if (!result) result = row.slice();
    result[override.col] = replacement;
  }
  return result ?? row;
}
var BrowserXlsbWriter = class {
  constructor(options = {}) {
    this._streamTableTotals = /* @__PURE__ */ new Map();
    this._useSharedStrings = options.useSharedStrings !== false;
    this._zip = new BrowserZip();
    this._sheetCount = 0;
    this._sheetList = [];
    this._sstDic = /* @__PURE__ */ new Map();
    this._sstRichDic = /* @__PURE__ */ new Map();
    this._sstValues = [];
    this._sstCntUnique = 0;
    this._sstCntAll = 0;
    this._nextTableId = 1;
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
    this._dxfs = new XlsbDxfRegistry((format) => this._styles.registerNumberFormat(format));
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
    return new Uint8Array(
      this._styles.toStylesBin(this._stylesBin, this._dxfs.toRecords())
    );
  }
  /** Resolve column widths; explicit widths win over the conservative autofit. */
  _resolveColumnWidths(columnCount, headers, rows, options, doAutofilter, headerStyle) {
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
        bold: headerStyle ? headerStyle.font?.bold === true : true,
        filterButton: doAutofilter
      });
    }
    const samples = options.sampleRows ?? rows;
    const limit = options.sampleRows ? samples?.length ?? 0 : resolveAutofitRows(options.autofitRows);
    if (samples && limit > 0) updateColWidthsFromRows(widths, samples, limit);
    return widths;
  }
  _presetHeaderStyle(options) {
    if (options.headerStyle && typeof options.headerStyle !== "string") {
      return mergeCellStyles({ font: { bold: true } }, options.headerStyle);
    }
    const preset = options.headerPreset ?? (typeof options.headerStyle === "string" ? options.headerStyle : void 0) ?? "bold";
    if (preset === "bold") return null;
    const base = {};
    if (preset === "bold+fill") base.font = { bold: true };
    if (preset === "fill" || preset === "bold+fill") {
      base.fill = { patternType: "solid", fgColor: "FFD9E1F2" };
    }
    return base;
  }
  _headerStyleIndex(layout) {
    return layout.headerStyle ? this._styles.register(layout.headerStyle) : 3;
  }
  _resolveLayout(options, hasHeaders, columnCount, headers, rowCount = 0) {
    const rowHeights = {};
    for (const [key, value] of Object.entries(options.rowHeights ?? {})) {
      const row = Number(key);
      if (!Number.isInteger(row) || row < 0) {
        throw new RangeError(`rowHeights keys must be non-negative row indices, got '${key}'`);
      }
      rowHeights[row] = validateRowHeight(value);
    }
    const hiddenRows = /* @__PURE__ */ new Set();
    for (const row of options.hiddenRows ?? []) {
      if (!Number.isInteger(row) || row < 0) {
        throw new RangeError(`hiddenRows must contain non-negative row indices, got ${row}`);
      }
      hiddenRows.add(row);
    }
    const hiddenColumns = [];
    for (const col of options.hiddenColumns ?? []) {
      if (!Number.isInteger(col) || col < 0 || col >= columnCount) {
        throw new RangeError(`hiddenColumns index ${col} is outside the sheet`);
      }
      hiddenColumns.push(col);
    }
    const rowGroups = normalizeGroups(options.rowGroups, "row");
    const columnGroups = normalizeGroups(options.columnGroups, "column");
    const rowMax = Math.max(
      rowCount - 1,
      ...rowGroups.map((group) => group.end),
      ...Object.keys(rowHeights).map(Number),
      ...hiddenRows,
      0
    );
    const columnMax = Math.max(
      columnCount - 1,
      ...columnGroups.map((group) => group.end),
      ...hiddenColumns,
      0
    );
    const tables = [];
    for (const spec of options.tables ?? []) {
      let columns = spec.columns;
      if ((!columns || columns.length === 0) && headers) {
        const range = parseRangeRef(spec.ref);
        columns = headers.slice(range.c1, range.c2 + 1);
      }
      tables.push(normalizeTableSpec({ ...spec, columns }, this._nextTableId++));
    }
    if (options.autoFilterRef !== void 0) parseRangeRef(options.autoFilterRef);
    return {
      headerStyle: this._presetHeaderStyle(options),
      styleRanges: new StyleRangeSet(options.styles),
      rowHeights,
      hiddenRows,
      hiddenColumns,
      rowOutline: buildOutline(rowGroups, rowMax),
      columnOutline: buildOutline(columnGroups, columnMax),
      freezePanes: resolveFreezePanes(options.freezePanes, hasHeaders),
      showGridLines: options.showGridLines !== false,
      showRowColHeaders: options.showRowColHeaders !== false,
      showZeroValues: options.showZeroValues !== false,
      rightToLeft: options.rightToLeft === true,
      zoom: validateZoom(options.zoom),
      tabColor: options.tabColor,
      sheetProtection: options.sheetProtection,
      autoFilterRef: options.autoFilterRef,
      tables
    };
  }
  _rowHeaderOptions(layout, rowIndex) {
    if (!layout) return {};
    const options = {};
    const height = layout.rowHeights[rowIndex];
    if (height !== void 0) options.heightTwips = Math.round(height * 20);
    if (layout.hiddenRows.has(rowIndex)) options.hidden = true;
    if (rowIndex < layout.rowOutline.level.length) {
      options.outlineLevel = layout.rowOutline.level[rowIndex];
      if (layout.rowOutline.hidden[rowIndex]) options.hidden = true;
      if (layout.rowOutline.collapsed[rowIndex]) options.collapsed = true;
    }
    return options;
  }
  /** Store sheet features and build the worksheet extra parts. */
  _applySheetFeatures(sheet, features) {
    if (features.merges) sheet.merges = features.merges;
    if (features.dataValidations) sheet.dataValidations = features.dataValidations;
    if (features.conditionalFormats && features.conditionalFormats.length > 0) {
      sheet.conditionalFormats = features.conditionalFormats;
      sheet.cfRecords = buildXlsbConditionalFormatRecords(
        features.conditionalFormats,
        this._dxfs,
        1
      );
    }
    let nextRelId = 2;
    if (features.hyperlinks && features.hyperlinks.length > 0) {
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
    const layout = sheet.layout;
    if (layout && layout.tables.length > 0) {
      sheet.tableParts = [];
      for (let index = 0; index < layout.tables.length; index++) {
        const table = layout.tables[index];
        const id = this._nextTableId - layout.tables.length + index;
        const path = `xl/tables/table${id}.bin`;
        const data = buildXlsbTablePart(table, id);
        sheet.tableParts.push({ id, path, data, rId: `rId${nextRelId++}` });
        sheet.extraParts = [...sheet.extraParts ?? [], { name: path, data }];
      }
    }
    if (features.comments && features.comments.length > 0) {
      sheet.comments = features.comments;
      sheet.commentRId = `rId${nextRelId++}`;
      sheet.commentVmlRId = `rId${nextRelId++}`;
      const commentPath = `xl/comments/comment${sheet.sheetId}.bin`;
      const vmlPath = `xl/drawings/commentsDrawing${sheet.sheetId}.vml`;
      sheet.extraParts = [
        ...sheet.extraParts ?? [],
        { name: commentPath, data: buildXlsbCommentsPart(features.comments, commentPath) },
        {
          name: vmlPath,
          data: new TextEncoder().encode(buildXlsbCommentsVml(features.comments))
        }
      ];
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
    if (sheet.cfRecords && sheet.cfRecords.length > 0) {
      bb.write(sheet.cfRecords);
    }
  }
  _writeSheetFooterRecords(bb, sheet, sheetHeader) {
    bb.write(sheetHeader.subarray(290, 346));
    const relIds = (sheet.tableParts ?? []).map((part) => part.rId);
    if (relIds.length > 0) bb.write(buildXlsbListParts(relIds));
    if (sheet.commentVmlRId !== void 0) {
      bb.write(buildXlsbLegacyDrawingRecord(sheet.commentVmlRId));
    }
    bb.write(sheetHeader.subarray(346));
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
    const { hidden = false, doAutofilter = true } = options;
    this.addSheet(sheetName, hidden);
    const sheet = this._sheetList[this._sheetCount - 1];
    const layout = this._resolveLayout(options, headers !== void 0, columnCount, headers);
    sheet.layout = layout;
    this._streamTableTotals = /* @__PURE__ */ new Map();
    for (const override of tableTotalsOverrides(layout.tables)) {
      const list = this._streamTableTotals.get(override.row) ?? [];
      list.push(override);
      this._streamTableTotals.set(override.row, list);
    }
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
      doAutofilter && headers !== void 0,
      layout.headerStyle ?? void 0
    );
    this._streamPending = new BrowserBigBuffer();
    this._streamHeadFlushed = false;
    this._streamHasHeaders = headers !== void 0;
    this._streamDataRows = 0;
    this._streamAutofitRows = resolveAutofitRows(options.autofitRows);
    if (headers) {
      const pending = this._streamPending;
      this.createRowHeader(
        pending,
        this._currentSheetRowNum,
        this._currentSheetStartCol,
        this._currentSheetEndCol,
        this._rowHeaderOptions(layout, 0)
      );
      const headerStyle = this._headerStyleIndex(layout);
      const baseHeader = layout.headerStyle ?? { font: { bold: true } };
      for (let c = 0; c < headers.length; c++) {
        const rangeStyle = layout.styleRanges.styleFor(0, c);
        const style = rangeStyle ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle)) : headerStyle;
        this.writeString(pending, headers[c], c, 0, style);
      }
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
    const layout = this._sheetList[this._sheetCount - 1]?.layout;
    const totals = this._streamTableTotals.get(this._currentSheetRowNum);
    if (totals) row = withTableTotals2(row, totals);
    this.createRowHeader(
      bb,
      this._currentSheetRowNum,
      this._currentSheetStartCol,
      this._currentSheetEndCol,
      this._rowHeaderOptions(layout, this._currentSheetRowNum)
    );
    for (let c = 0; c < row.length; c++) {
      this._writeCell(
        bb,
        row[c],
        c,
        layout?.styleRanges.styleFor(this._currentSheetRowNum, c)
      );
    }
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
    const sheet = this._sheetList[this._sheetCount - 1];
    const sheetHeader = new Uint8Array(this._sheet1Bytes);
    applyXlsbSheetProtection(sheetHeader, sheet.layout?.sheetProtection);
    bb.write(sheetHeader.subarray(218, 290));
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
      sheet.filterData = {
        startRow: 0,
        endRow: this._currentSheetRowNum - 1,
        startColumn: this._currentSheetStartCol,
        endColumn: this._currentSheetEndCol - 1
      };
    }
    this._writeFeatureRecords(bb, sheet);
    this._writeSheetFooterRecords(bb, sheet, sheetHeader);
    bb.patchInt32LE(36, Math.max(0, this._currentSheetRowNum - 1));
    this._zip.addFile(sheet.pathInArchive, bb.toUint8Array());
    this._isStreaming = false;
    this._currentSheetBuffer = null;
    this._streamPending = null;
    this._streamHeadFlushed = true;
    this._streamHasHeaders = false;
    this._streamDataRows = 0;
    this._streamAutofitRows = AUTOFIT_SAMPLE_ROWS;
    this._streamTableTotals = /* @__PURE__ */ new Map();
  }
  writeSheet(rows, headers = null, options = {}) {
    const doAutofilter = options.doAutofilter !== false;
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
    const layout = this._resolveLayout(
      options,
      headers !== null,
      colCount,
      headers,
      totalBr
    );
    sheet.layout = layout;
    this._applySheetFeatures(sheet, options);
    const totalsOverrides = tableTotalsOverrides(layout.tables);
    if (totalsOverrides.length > 0) {
      const headerOffset = headers ? 1 : 0;
      const byRow = /* @__PURE__ */ new Map();
      for (const override of totalsOverrides) {
        const index = override.row - headerOffset;
        if (index < 0 || index >= rows.length) {
          throw new Error(
            `Table totals row ${override.row + 1} is outside the written rows (the sheet has ${totalBr} rows)`
          );
        }
        const list = byRow.get(index) ?? [];
        list.push(override);
        byRow.set(index, list);
      }
      for (const [index, list] of byRow) {
        rows[index] = withTableTotals2(rows[index], list);
      }
    }
    this._colWidths = this._resolveColumnWidths(
      colCount,
      headers,
      rows,
      options,
      doAutofilter && headers !== null,
      layout.headerStyle ?? void 0
    );
    const doAf = doAutofilter && !!headers;
    this._writeSheetHeader(bb, 0, colCount, doAf);
    let rn = 0;
    if (headers) {
      this.createRowHeader(bb, rn, 0, colCount, this._rowHeaderOptions(layout, 0));
      const headerStyle = this._headerStyleIndex(layout);
      const baseHeader = layout.headerStyle ?? { font: { bold: true } };
      for (let c = 0; c < headers.length; c++) {
        const rangeStyle = layout.styleRanges.styleFor(0, c);
        const style = rangeStyle ? this._styles.register(mergeCellStyles(baseHeader, rangeStyle)) : headerStyle;
        this.writeString(bb, headers[c], c, 0, style);
      }
      rn++;
    }
    for (let r = 0; r < rows.length; r++) {
      this.createRowHeader(bb, rn, 0, colCount, this._rowHeaderOptions(layout, rn));
      for (let c = 0; c < rows[r].length; c++) {
        this._writeCell(bb, rows[r][c], c, layout.styleRanges.styleFor(rn, c));
      }
      rn++;
    }
    const sheetHeader = new Uint8Array(this._sheet1Bytes);
    applyXlsbSheetProtection(sheetHeader, layout.sheetProtection);
    bb.write(sheetHeader.subarray(218, 290));
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
      sheet.filterData = {
        startRow: 0,
        endRow: rn - 1,
        startColumn: 0,
        endColumn: colCount - 1
      };
    }
    this._writeFeatureRecords(bb, sheet);
    this._writeSheetFooterRecords(bb, sheet, sheetHeader);
    bb.patchInt32LE(36, Math.max(0, rn - 1));
    this._zip.addFile(sheet.pathInArchive, bb.toUint8Array());
  }
  _writeSheetHeader(bb, startCol, endCol, _doAf = false) {
    const layout = this._sheetList[this._sheetCount - 1]?.layout;
    const sh = new Uint8Array(this._sheet1Bytes);
    applyXlsbSheetProtection(sh, layout?.sheetProtection);
    const shView = new DataView(sh.buffer);
    shView.setInt32(40, startCol, true);
    shView.setInt32(44, endCol - 1, true);
    if (layout?.tabColor !== void 0) {
      sh.set(xlsbColorPayload(layout.tabColor), 9);
    }
    let viewFlags = this._sheetCount === 1 ? 220 : 156;
    if (layout?.showGridLines === false) viewFlags &= ~4;
    if (layout?.showRowColHeaders === false) viewFlags &= ~8;
    if (layout?.showZeroValues === false) viewFlags &= ~16;
    if (layout?.rightToLeft) viewFlags |= 32;
    sh[54] = viewFlags;
    bb.write(sh.subarray(0, 55));
    let viewTail = sh.subarray(55, 84);
    if (layout?.zoom !== void 0) {
      const scale = new Uint8Array(4);
      const scaleView = new DataView(scale.buffer);
      scaleView.setUint16(0, layout.zoom, true);
      scaleView.setUint16(2, layout.zoom, true);
      const merged = new Uint8Array(viewTail.length);
      merged.set(viewTail.subarray(0, 17), 0);
      merged.set(scale, 17);
      merged.set(viewTail.subarray(21), 21);
      viewTail = merged;
    }
    bb.write(viewTail);
    bb.write(buildXlsbPaneRecord(layout?.freezePanes ?? null));
    bb.write(buildXlsbSelectionRecord(layout?.freezePanes ?? null));
    bb.write(sh.subarray(123, 157));
    bb.write(new Uint8Array([
      layout?.rowOutline.maxLevel ?? 0,
      layout?.columnOutline.maxLevel ?? 0
    ]));
    const widths = this._colWidths.map((width) => defaultColWidth(width));
    bb.write(buildXlsbColumnRecords({
      startCol,
      endCol,
      widths,
      hiddenColumns: layout?.hiddenColumns ?? [],
      outline: layout?.columnOutline
    }));
    bb.write(sh.subarray(159, 175));
    bb.write(new Uint8Array([38, 0]));
  }
  _writeCell(bb, raw, c, extraStyle) {
    if (raw === null || raw === void 0) return;
    const cellStyle = getCellStyle(raw);
    const style = extraStyle ? mergeCellStyles(cellStyle, extraStyle) : cellStyle;
    let styleNum = style !== null ? this._styles.register(style) : 0;
    const formula = getFormula(raw);
    if (formula !== null) {
      const rawCached = raw.value;
      let cached;
      if (rawCached instanceof Date) {
        cached = timeToOaDate(rawCached.getTime(), this._oaEpoch);
        if (styleNum === 0) styleNum = dateHasTime(rawCached) ? 1 : 2;
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
      this._writeRichText(bb, val, c, styleNum);
      return;
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
      const dateStyle = dateHasTime(val) ? 1 : 2;
      this.writeDouble(bb, oaDate, c, styleNum !== 0 ? styleNum : dateStyle);
    } else {
      this.writeString(bb, val.toString(), c, 0, styleNum);
    }
  }
  createRowHeader(bb, rn, sc, ec, options = {}) {
    writeXlsbRowHeader(bb, rn, sc, ec, options);
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
    this.writeDouble(bb, timeToOaDate(date.getTime(), this._oaEpoch), colNum, dateHasTime(date) ? 1 : 2);
  }
  _writeRichText(bb, value, colNum, styleNum) {
    if (!this._useSharedStrings) {
      this.writeString(bb, value.toString(), colNum, 0, styleNum);
      return;
    }
    const runs = [];
    let offset = 0;
    for (const run of value.runs) {
      runs.push({ ich: offset, ifnt: this._styles.registerRichFont(run.style?.font) });
      offset += run.text.length;
    }
    const text = value.toString();
    assertCellTextLength(text);
    const key = JSON.stringify({ text, runs });
    let index = this._sstRichDic.get(key);
    if (index === void 0) {
      index = this._sstValues.length;
      this._sstRichDic.set(key, index);
      this._sstValues.push({ text, runs });
      this._sstCntUnique++;
    }
    this._sstCntAll++;
    writeXlsbIsst(bb, index, colNum, styleNum);
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
      index = this._sstValues.length;
      this._sstDic.set(val, index);
      this._sstValues.push(val);
      this._sstCntUnique++;
    }
    this._sstCntAll++;
    writeXlsbIsst(bb, index, colNum, finalStyle);
  }
  _saveSst() {
    const bb = new BrowserBigBuffer();
    writeXlsbSharedStrings(bb, this._sstValues, this._sstCntAll, this._sstCntUnique);
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
    for (const sheet of this._sheetList) {
      for (const part of sheet.extraParts ?? []) {
        this._zip.addFile(part.name, new Uint8Array(part.data));
      }
    }
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
    const hasComments = this._sheetList.some((sheet) => (sheet.comments?.length ?? 0) > 0);
    let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`;
    if (hasComments) {
      ct += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
    }
    for (const s of this._sheetList) {
      ct += `<Override PartName="/${s.pathInArchive}" ContentType="application/vnd.ms-excel.worksheet"/>`;
      ct += `<Override PartName="/xl/worksheets/binaryIndex${s.sheetId}.bin" ContentType="application/vnd.ms-excel.binIndexWs"/>`;
      for (const part of s.tableParts ?? []) {
        ct += `<Override PartName="/${part.path}" ContentType="application/vnd.ms-excel.table"/>`;
      }
      if ((s.comments?.length ?? 0) > 0) {
        ct += `<Override PartName="/xl/comments/comment${s.sheetId}.bin" ContentType="application/vnd.ms-excel.comments"/>`;
      }
    }
    ct += `<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>
<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
    this._zip.addFile("[Content_Types].xml", ct);
    this._zip.addFile("xl/theme/theme1.xml", XLSX_THEME_XML);
    let wr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const s of this._sheetList) wr += `<Relationship Id="rId${s.sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${s.nameInArchive}"/>`;
    wr += `<Relationship Id="rId${this._sheetList.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>
<Relationship Id="rId${this._sheetList.length + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
<Relationship Id="rId${this._sheetList.length + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
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
      for (const part of s.tableParts ?? []) {
        ws += `<Relationship Id="${part.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/${part.path.split("/").pop()}"/>`;
      }
      if (s.commentRId !== void 0) {
        ws += `<Relationship Id="${s.commentRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment${s.sheetId}.bin"/>`;
      }
      if (s.commentVmlRId !== void 0) {
        ws += `<Relationship Id="${s.commentVmlRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/commentsDrawing${s.sheetId}.vml"/>`;
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
  RichText,
  downloadXlsb,
  downloadXlsbMultiSheet,
  downloadXlsx,
  downloadXlsxMultiSheet,
  mergeCellStyles
};
