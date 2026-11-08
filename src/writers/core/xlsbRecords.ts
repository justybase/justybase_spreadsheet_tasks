/**
 * Shared BIFF12 (XLSB) record serialization.
 *
 * Both the Node and the browser XLSB writers use these functions, so a cell
 * written in Node and the same cell written in a browser produce identical
 * bytes. The module has no platform imports.
 */
import { ByteSink } from './ByteSink';
import { excelErrorToBiffCode } from '../../biff12Utils';

function decodeBase64(base64: string): Uint8Array {
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
}

/**
 * Base `xl/styles.bin`: 2 fonts, 2 fills, 1 border, 6 cell XFs
 * (normal, date, datetime, bold, fill, bold+fill), and the two custom date
 * formats (164, 166). Custom cell formats are appended by
 * {@link buildXlsbStylesBin}.
 */
export const XLSB_STYLES_TEMPLATE: Uint8Array = decodeBase64(
    'lgIA5wQEAgAAACwspAATAAAAeQB5AHkAeQBcAC0AbQBtAFwALQBkAGQAXAAgAGgAaAA6AG0AbQAsHqYADAAAAHkAeQB5AHkAXAAtAG0AbQBcAC0AZABkAOgEAOMEBAIAAAArJ9wAAACQAQAAAAIAAAcBAAAAAAD/AgcAAABDAGEAbABpAGIAcgBpACsn3AABALwCAAAAAu4ABwEAAAAAAP8CBwAAAEMAYQBsAGkAYgByAGkAJQYBAAIOAICBCAAmAOQEANsEBAIAAAAtRAAAAAADQAAAAAAA/wNBAAD/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALUQRAAAAA0AAAAAAAP8DQQAA/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANwEAOUEBAEAAAAuMwAAAAEAAAAAAAAAAAABAAAAAAAAAAAAAQAAAAAAAAAAAAEAAAAAAAAAAAABAAAAAAAAAOYEAPIEBAEAAAAvEP//AAAAAAAAAAAAABAQAADzBADpBAQGAAAALxAAAAAAAAAAAAAAAAAQEAAALxAAAKQAAAAAAAAAAAAQEAEALxAAAKYAAAAAAAAAAAAQEAEALxAAAAAAAQAAAAAAAAAQEAEALxAAAAAAAAABAAAAAAAQEAEALxAAAAAAAQABAAAAAAAQEAEA6gQA6wQEAQAAACUGAQACEQCAgBgQAAAAAAAAAAAAAAAAAAAAACYAMBwAAAAAAQAAAAgAAABOAG8AcgBtAGEAbABuAHkA7AQA+QMEAAAAAPoDAPwDUAAAAAARAAAAVABhAGIAbABlAFMAdAB5AGwAZQBNAGUAZABpAHUAbQAyABEAAABQAGkAdgBvAHQAUwB0AHkAbABlAEwAaQBnAGgAdAAxADYA/QMAIwQCDgAA6wgA9ggqAAAAABEAAABTAGwAaQBjAGUAcgBTAHQAeQBsAGUATABpAGcAaAB0ADEA9wgA7AgAJAAjBAMPAACwEACyEDIAAAAAFQAAAFQAaQBtAGUAUwBsAGkAYwBlAHIAUwB0AHkAbABlAEwAaQBnAGgAdAAxALMQALEQACQAlwIA',
);

/** Cell XFs that always exist in {@link XLSB_STYLES_TEMPLATE}. */
export const XLSB_BASE_XF_COUNT = 6;

export function concat(chunks: Uint8Array[]): Uint8Array {
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

function encodeVlq(value: number): number[] {
    const result: number[] = [];
    while (value >= 0x80) {
        result.push((value & 0x7f) | 0x80);
        value >>>= 7;
    }
    result.push(value & 0x7f);
    return result;
}

export function buildBrtRecord(type: number, data: Uint8Array): Uint8Array {
    const header = [...encodeVlq(type), ...encodeVlq(data.length)];
    return concat([new Uint8Array(header), data]);
}

function buildStFmtRecord(ifmt: number, formatString: string): Uint8Array {
    const cch = formatString.length;
    const remaining = 2 + 4 + cch * 2;
    const data = new Uint8Array(2 + remaining);
    data[0] = 0x2c;
    data[1] = remaining;
    const view = new DataView(data.buffer);
    view.setUint16(2, ifmt, true);
    view.setUint32(4, cch, true);
    for (let i = 0; i < cch; i++) {
        const code = formatString.charCodeAt(i);
        data[8 + i * 2] = code & 0xff;
        data[9 + i * 2] = (code >> 8) & 0xff;
    }
    return data;
}

function buildXfRecord(fontId: number, ifmt: number, flags: number, byte6Extra: number = 0): Uint8Array {
    const data = new Uint8Array(18);
    data[0] = 0x2f;
    data[1] = 16;
    const view = new DataView(data.buffer);
    view.setUint16(2, fontId, true);
    view.setUint16(4, ifmt, true);
    data[6] = byte6Extra;
    data[14] = 0x10;
    data[15] = 0x10;
    view.setUint16(16, flags, true);
    return data;
}

export interface XlsbFormatRecord {
    numfmtId: number;
    xfIndex: number;
}

/**
 * Build `xl/styles.bin`, appending custom number formats after the six base
 * XFs. Returns the template untouched when there are no custom formats.
 */
export function buildXlsbStylesBin(
    template: Uint8Array,
    formats: Map<string, XlsbFormatRecord>,
): Uint8Array {
    if (formats.size === 0) return template;

    const base = template;

    let fontEndOff = -1;
    for (let i = 3; i < base.length - 3; i++) {
        if (base[i] === 0xe8 && base[i + 1] === 0x04 && base[i + 2] === 0x00) {
            fontEndOff = i;
            break;
        }
    }

    let xfBeginOff = -1;
    for (let i = fontEndOff + 3; i < base.length - 2; i++) {
        if (base[i] === 0xe9 && base[i + 1] === 0x04) {
            xfBeginOff = i;
            break;
        }
    }

    let xfEndOff = -1;
    for (let i = xfBeginOff; i < base.length - 3; i++) {
        if (base[i] === 0xea && base[i + 1] === 0x04 && base[i + 2] === 0x00) {
            xfEndOff = i;
            break;
        }
    }

    let fillBeginOff = -1;
    for (let i = fontEndOff + 3; i < base.length - 2; i++) {
        if (base[i] === 0xe3 && base[i + 1] === 0x04) {
            fillBeginOff = i;
            break;
        }
    }

    const xfBeginBody = xfBeginOff + 3 + 4; // record header + uint16 count + 2 pad bytes
    const baseXfRecords = base.subarray(xfBeginBody, xfEndOff);
    const totalFmt = 2 + formats.size;
    const totalXf = XLSB_BASE_XF_COUNT + formats.size;

    const parts: Uint8Array[] = [];

    parts.push(base.subarray(0, 3));

    const fmtHeader = new Uint8Array(4);
    new DataView(fmtHeader.buffer).setUint16(0, totalFmt, true);
    parts.push(buildBrtRecord(0x0267, fmtHeader));

    parts.push(buildStFmtRecord(164, 'yyyy\\-mm\\-dd\\ hh:mm'));
    parts.push(buildStFmtRecord(166, 'yyyy\\-mm\\-dd'));
    for (const [formatString, record] of formats) {
        parts.push(buildStFmtRecord(record.numfmtId, formatString));
    }

    parts.push(buildBrtRecord(0x0268, new Uint8Array(0)));
    parts.push(base.subarray(fillBeginOff, xfBeginOff));

    const xfHeader = new Uint8Array(4);
    new DataView(xfHeader.buffer).setUint16(0, totalXf, true);
    parts.push(buildBrtRecord(0x0269, xfHeader));

    parts.push(baseXfRecords);

    const sorted = [...formats.entries()].sort((a, b) => a[1].xfIndex - b[1].xfIndex);
    for (const [, record] of sorted) {
        parts.push(buildXfRecord(0, record.numfmtId, 0x0001));
    }

    parts.push(buildBrtRecord(0x026a, new Uint8Array(0)));
    parts.push(base.subarray(xfEndOff + 3));

    return concat(parts);
}

/** BrtRowHdr (0x00), 27-byte payload. */
export function writeXlsbRowHeader(sink: ByteSink, rowNum: number, startCol: number, endCol: number): void {
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

/** BrtCellRk (0x02) holding an integer. */
export function writeXlsbRkInteger(sink: ByteSink, val: number, colNum: number, styleNum: number = 0): void {
    sink.ensureCapacity(14);
    sink.writeUnsafeByte(2);
    sink.writeUnsafeByte(12);
    sink.writeUnsafeInt32LE(colNum);
    sink.writeUnsafeByte(styleNum);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeInt32LE((val << 2) | 2);
}

/** BrtCellReal (0x05). */
export function writeXlsbDouble(sink: ByteSink, val: number, colNum: number, styleNum: number = 0): void {
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

/** BrtCellBool (0x04). */
export function writeXlsbBool(sink: ByteSink, val: boolean, colNum: number, styleNum: number = 0): void {
    sink.ensureCapacity(13);
    sink.writeUnsafeByte(0x04);
    sink.writeUnsafeByte(9);
    sink.writeUnsafeInt32LE(colNum);
    sink.writeUnsafeByte(styleNum);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(val ? 1 : 0);
}

/** BrtCellError (0x03): column, 24-bit style and the BIFF12 error code. */
export function writeXlsbError(sink: ByteSink, code: string, colNum: number, styleNum: number = 0): void {
    sink.ensureCapacity(13);
    sink.writeUnsafeByte(0x03);
    sink.writeUnsafeByte(9);
    sink.writeUnsafeInt32LE(colNum);
    sink.writeUnsafeByte(styleNum);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(0);
    sink.writeUnsafeByte(excelErrorToBiffCode(code));
}

/** BrtCellIsst (0x07): shared-string index. */
export function writeXlsbIsst(sink: ByteSink, index: number, colNum: number, styleNum: number = 0): void {
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

/** BrtCellSt (0x06): inline UTF-16 string. */
export function writeXlsbInlineString(sink: ByteSink, value: string, colNum: number, styleNum: number = 0): void {
    const cch = value.length;
    const len = 12 + cch * 2;
    sink.ensureCapacity(4 + len);
    sink.writeUnsafeByte(6);
    let length = len;
    while (length >= 0x80) {
        sink.writeUnsafeByte((length & 0x7f) | 0x80);
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

/**
 * Write the BrtBeginSst / BrtBundleSh-equivalent SST part. Kept byte-compatible
 * with the previous per-writer implementations.
 */
export function writeXlsbSharedStrings(
    sink: ByteSink,
    entries: Iterable<string>,
    countAll: number,
    countUnique: number,
): void {
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
            sink.writeByte(128 + (recordLength % 128));
            const high = recordLength >> 7;
            if (high >= 256) {
                sink.writeByte(128 + (high % 128));
            } else {
                sink.writeByte(high);
            }
            sink.writeByte(recordLength >> 14);
            if ((recordLength >> 14) > 0) sink.writeByte(0);
        } else {
            sink.writeByte(recordLength & 0xff);
            sink.writeByte((recordLength >> 8) & 0xff);
        }

        sink.writeInt32LE(textLength);
        sink.writeUtf16LE(text);
    }

    sink.writeByte(160);
    sink.writeByte(1);
    sink.writeByte(0);
}
