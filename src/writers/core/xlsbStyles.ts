/**
 * XLSB cell style registry.
 *
 * Builds the BIFF12 records for fonts, fills, borders, number formats and cell
 * XFs, then rewrites the `xl/styles.bin` template with the registered custom
 * components appended after the six base XFs. Shared by the Node and browser
 * XLSB writers, so both produce identical bytes for the same styles.
 */
import {
    BorderSide,
    BorderStyle,
    CellStyle,
    FillStyle,
    FontStyle,
    StyleColor,
} from '../../Formats';
import { normalizeStyle } from './xlsxParts';
import { buildBrtRecord, concat, XLSB_BASE_XF_COUNT } from './xlsbRecords';

const BRT_FONT = 0x002b;
const BRT_FMT = 0x002c;
const BRT_FILL = 0x002d;
const BRT_BORDER = 0x002e;
const BRT_XF = 0x002f;

const BRT_BEGIN_FMTS = 0x0267;
const BRT_END_FMTS = 0x0268;
const BRT_BEGIN_FONTS = 0x0263;
const BRT_END_FONTS = 0x0264;
const BRT_BEGIN_FILLS = 0x025b;
const BRT_END_FILLS = 0x025c;
const BRT_BEGIN_BORDERS = 0x0265;
const BRT_END_BORDERS = 0x0266;
const BRT_BEGIN_CELL_XFS = 0x0269;
const BRT_END_CELL_XFS = 0x026a;
const BRT_BEGIN_CELL_STYLE_XFS = 0x0272;
const BRT_END_CELL_STYLE_XFS = 0x0273;

const UNDERLINE_BYTES: Record<string, number> = {
    none: 0x00,
    single: 0x01,
    double: 0x02,
    singleAccounting: 0x21,
    doubleAccounting: 0x22,
};

const BASELINE_BYTES: Record<string, number> = {
    baseline: 0,
    superscript: 1,
    subscript: 2,
};

const PATTERN_BYTES: Record<string, number> = {
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
    gray0625: 18,
};

const BORDER_BYTES: Record<string, number> = {
    none: 0x00,
    thin: 0x01,
    medium: 0x02,
    dashed: 0x03,
    dotted: 0x04,
    thick: 0x05,
    double: 0x06,
    hair: 0x07,
    mediumDashed: 0x08,
    dashDot: 0x09,
    mediumDashDot: 0x0a,
    dashDotDot: 0x0b,
    mediumDashDotDot: 0x0c,
    slantDashDot: 0x0d,
};

const HORIZONTAL_BYTES: Record<string, number> = {
    general: 0,
    left: 1,
    center: 2,
    right: 3,
    fill: 4,
    justify: 5,
    centerContinuous: 6,
    distributed: 7,
};

const VERTICAL_BYTES: Record<string, number> = {
    top: 0,
    center: 1,
    bottom: 2,
    justify: 3,
    distributed: 4,
};

function u16(value: number): number[] {
    const v = value & 0xffff;
    return [v & 0xff, (v >>> 8) & 0xff];
}

function u32(value: number): number[] {
    const v = value >>> 0;
    return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
}

function utf16leBytes(text: string): number[] {
    const out: number[] = [];
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        out.push(code & 0xff, (code >>> 8) & 0xff);
    }
    return out;
}

/** BrtColor ([MS-XLSB] 2.4.334); all zeroes mean automatic. */
function colorBytes(color: StyleColor | undefined): number[] {
    if (color === undefined) return [0, 0, 0, 0, 0, 0, 0, 0];
    if (typeof color === 'string') {
        let value = color.trim().replace(/^#/, '').toUpperCase();
        if (value.length === 6) value = `FF${value}`;
        if (!/^[0-9A-F]{8}$/.test(value)) {
            throw new Error(`Invalid color '${color}'. Use RRGGBB, AARRGGBB or '#' + hex.`);
        }
        const alpha = parseInt(value.slice(0, 2), 16);
        const red = parseInt(value.slice(2, 4), 16);
        const green = parseInt(value.slice(4, 6), 16);
        const blue = parseInt(value.slice(6, 8), 16);
        return [0x05, 0, 0, 0, red, green, blue, alpha];
    }
    if ('indexed' in color) {
        return [0x03, color.indexed & 0xff, 0, 0, 0, 0, 0, 0xff];
    }
    const tint = Math.max(-32768, Math.min(32767, Math.round((color.tint ?? 0) * 32767)));
    return [0x07, color.theme & 0xff, ...u16(tint), 0, 0, 0, 0xff];
}

/** BrtFont ([MS-XLSB] 2.4.683). */
export function buildXlsbFontRecord(font: FontStyle): Uint8Array {
    if (font.name !== undefined && font.name.length > 31) {
        throw new RangeError('XLSB font names are limited to 31 characters');
    }
    let flags = 0;
    if (font.bold) flags |= 0x01;
    if (font.italic) flags |= 0x02;
    if (font.strike) flags |= 0x08;
    if (font.outline) flags |= 0x10;
    if (font.shadow) flags |= 0x20;
    if (font.condense) flags |= 0x40;
    if (font.extend) flags |= 0x80;

    const height = Math.max(0x14, Math.min(0x1fff, Math.round((font.size ?? 11) * 20)));
    const baseline = BASELINE_BYTES[font.vertAlign ?? 'baseline'] ?? 0;
    const underline = font.underline === true
        ? UNDERLINE_BYTES.single
        : typeof font.underline === 'string'
            ? UNDERLINE_BYTES[font.underline] ?? 0
            : 0;

    const payload: number[] = [
        ...u16(height),
        ...u16(flags),
        ...u16(font.bold ? 700 : 400),
        ...u16(baseline),
        underline,
        font.family ?? 2,
        font.charset ?? 0,
        0,
        ...colorBytes(font.color),
    ];
    let scheme = font.scheme === 'major' ? 1 : font.scheme === 'minor' ? 2 : 0;
    if (font.scheme === undefined && font.name === undefined) {
        // Default Calibri is the minor theme font; a named font without an
        // explicit scheme must not claim one.
        scheme = 2;
    }
    payload.push(scheme);
    const name = utf16leBytes(font.name ?? 'Calibri');
    payload.push(...u32(name.length / 2), ...name);
    return buildBrtRecord(BRT_FONT, Uint8Array.from(payload));
}

/** BrtFill ([MS-XLSB] 2.4.674). */
export function buildXlsbFillRecord(fill: FillStyle): Uint8Array {
    const payload: number[] = [
        ...u32(PATTERN_BYTES[fill.patternType ?? 'none'] ?? 0),
        ...colorBytes(fill.fgColor),
        ...colorBytes(fill.bgColor),
    ];
    while (payload.length < 68) payload.push(0);
    return buildBrtRecord(BRT_FILL, Uint8Array.from(payload));
}

function borderSide(side: BorderSide | undefined): number[] {
    const style = side?.style !== undefined ? (BORDER_BYTES[side.style] ?? 0) : 0;
    return [style, 0, ...colorBytes(side?.color)];
}

/** BrtBorder ([MS-XLSB] 2.4.311). */
export function buildXlsbBorderRecord(border: BorderStyle): Uint8Array {
    let diagonal = 0;
    if (border.diagonalDown) diagonal |= 0x01;
    if (border.diagonalUp) diagonal |= 0x02;
    const payload: number[] = [
        diagonal,
        ...borderSide(border.top),
        ...borderSide(border.bottom),
        ...borderSide(border.left),
        ...borderSide(border.right),
        ...borderSide(border.diagonal),
    ];
    return buildBrtRecord(BRT_BORDER, Uint8Array.from(payload));
}

interface XlsbXfOptions {
    ifmt?: number;
    fontId?: number;
    fillId?: number;
    borderId?: number;
    style?: CellStyle;
}

/** BrtXF ([MS-XLSB] 2.4.863). */
export function buildXlsbXfRecord(options: XlsbXfOptions): Uint8Array {
    const { ifmt = 0, fontId = 0, fillId = 0, borderId = 0, style } = options;
    const alignment = style?.alignment;
    const protection = style?.protection;

    let rotation = 0;
    let indent = 0;
    let alignmentByte = 0x10;
    let flagsByte = 0x10;
    if (alignment) {
        rotation = alignment.textRotation ?? 0;
        indent = alignment.indent ?? 0;
        alignmentByte = (HORIZONTAL_BYTES[alignment.horizontal ?? 'general'] ?? 0)
            | ((VERTICAL_BYTES[alignment.vertical ?? 'bottom'] ?? 2) << 3)
            | (alignment.wrapText ? 0x40 : 0)
            | (alignment.justifyLastLine ? 0x80 : 0);
        flagsByte = 0x10
            | (alignment.shrinkToFit ? 0x01 : 0)
            | ((alignment.readingOrder ?? 0) << 2);
    }
    if (protection) {
        flagsByte = (flagsByte & 0xcf)
            | (protection.locked !== false ? 0x10 : 0)
            | (protection.hidden ? 0x20 : 0);
    }

    let flags = 0;
    if (ifmt !== 0) flags |= 0x01;
    if (fontId !== 0) flags |= 0x02;
    if (alignment) flags |= 0x04;
    if (borderId !== 0) flags |= 0x08;
    if (fillId !== 0) flags |= 0x10;
    if (protection) flags |= 0x20;

    const payload: number[] = [
        0, 0, // ixfeParent: built-in cell style XF 0
        ...u16(ifmt),
        ...u16(fontId),
        ...u16(fillId),
        ...u16(borderId),
        rotation & 0xff,
        indent & 0xff,
        alignmentByte & 0xff,
        flagsByte & 0xff,
        ...u16(flags),
    ];
    return buildBrtRecord(BRT_XF, Uint8Array.from(payload));
}

function numFmtRecord(ifmt: number, format: string): Uint8Array {
    if (format.length > 255) {
        throw new RangeError('XLSB number formats are limited to 255 characters');
    }
    const encoded = utf16leBytes(format);
    const payload = [...u16(ifmt), ...u32(encoded.length / 2), ...encoded];
    return buildBrtRecord(BRT_FMT, Uint8Array.from(payload));
}

interface Biff12RecordSpan {
    id: number;
    start: number;
    dataStart: number;
    end: number;
}

function scanRecords(data: Uint8Array): Biff12RecordSpan[] {
    const spans: Biff12RecordSpan[] = [];
    let pos = 0;
    while (pos < data.length) {
        const start = pos;
        let id = 0;
        let shift = 0;
        let byte: number;
        do {
            if (pos >= data.length) return spans;
            byte = data[pos++];
            id |= (byte & 0x7f) << shift;
            shift += 7;
        } while (byte & 0x80);
        let len = 0;
        shift = 0;
        do {
            if (pos >= data.length) return spans;
            byte = data[pos++];
            len |= (byte & 0x7f) << shift;
            shift += 7;
        } while (byte & 0x80);
        if (pos + len > data.length) return spans;
        spans.push({ id, start, dataStart: pos, end: pos + len });
        pos += len;
    }
    return spans;
}

function countRecord(id: number, count: number): Uint8Array {
    return buildBrtRecord(id, Uint8Array.from(u32(count)));
}

/**
 * Deduplicating style registry for `xl/styles.bin`. Base components stay at
 * their template ids: fonts 0-1, fills 0-1, border 0, cell XFs 0-5.
 */
export class XlsbStyleRegistry {
    private readonly fontIds = new Map<string, number>();
    private readonly fillIds = new Map<string, number>();
    private readonly borderIds = new Map<string, number>();
    private readonly xfIds = new Map<string, number>();
    private readonly formats = new Map<string, number>();

    private readonly fonts: Uint8Array[] = [];
    private readonly fills: Uint8Array[] = [];
    private readonly borders: Uint8Array[] = [];
    private readonly xfs: Uint8Array[] = [];
    private readonly formatRecords: Array<{ ifmt: number; format: string }> = [];

    private nextNumFmtId = 167;
    private nextFontId = 2;
    private nextFillId = 2;
    private nextBorderId = 1;

    /** Cell XF index for a style; 0 for null/empty styles. */
    register(style: CellStyle | null | undefined): number {
        const normalized = normalizeStyle(style);
        if (normalized === null) return 0;
        const key = JSON.stringify(normalized);
        const existing = this.xfIds.get(key);
        if (existing !== undefined) return existing;

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
            () => this.nextFontId++,
        );
        const fillId = this.componentId(
            this.fillIds,
            normalized.fill ? JSON.stringify(normalized.fill) : null,
            normalized.fill ? buildXlsbFillRecord(normalized.fill) : null,
            this.fills,
            null,
            () => this.nextFillId++,
        );
        const borderId = this.componentId(
            this.borderIds,
            normalized.border ? JSON.stringify(normalized.border) : null,
            normalized.border ? buildXlsbBorderRecord(normalized.border) : null,
            this.borders,
            null,
            () => this.nextBorderId++,
        );

        let ifmt = 0;
        if (normalized.numFmt !== undefined) {
            ifmt = this.numFmtIdFor(normalized.numFmt);
        }

        const index = XLSB_BASE_XF_COUNT + this.xfs.length;
        this.xfs.push(buildXlsbXfRecord({
            ifmt,
            fontId,
            fillId,
            borderId,
            style: normalized,
        }));
        this.xfIds.set(key, index);
        return index;
    }

    /** XF index for a plain number format (legacy FormattedCell path). */
    registerFormat(format: string): number {
        return this.register({ numFmt: format });
    }

    private numFmtIdFor(format: string): number {
        let id = this.formats.get(format);
        if (id === undefined) {
            if (this.nextNumFmtId > 382) {
                throw new RangeError('XLSB supports at most 216 custom number formats');
            }
            id = this.nextNumFmtId++;
            this.formats.set(format, id);
            this.formatRecords.push({ ifmt: id, format });
        }
        return id;
    }

    private componentId(
        registry: Map<string, number>,
        key: string | null,
        record: Uint8Array | null,
        target: Uint8Array[],
        baseId: number | null,
        allocate: () => number,
    ): number {
        if (key === null) return 0;
        const existing = registry.get(key);
        if (existing !== undefined) return existing;
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
    toStylesBin(template: Uint8Array): Uint8Array {
        if (this.xfs.length === 0 && this.fonts.length === 0
            && this.fills.length === 0 && this.borders.length === 0
            && this.formatRecords.length === 0) {
            return template;
        }

        const records = scanRecords(template);
        const find = (id: number): Biff12RecordSpan => {
            const record = records.find(r => r.id === id);
            if (!record) throw new Error(`XLSB styles template is missing record 0x${id.toString(16)}`);
            return record;
        };
        const section = (beginId: number, endId: number) => {
            const begin = find(beginId);
            const end = find(endId);
            const inner = records.filter(r => r.start >= begin.end && r.end <= end.start);
            return { begin, end, inner };
        };

        const parts: Uint8Array[] = [];
        const fmts = section(BRT_BEGIN_FMTS, BRT_END_FMTS);
        parts.push(template.subarray(0, fmts.begin.start));
        parts.push(countRecord(BRT_BEGIN_FMTS, fmts.inner.length + this.formatRecords.length));
        for (const record of fmts.inner) parts.push(template.subarray(record.start, record.end));
        for (const fmt of this.formatRecords) parts.push(numFmtRecord(fmt.ifmt, fmt.format));
        parts.push(template.subarray(fmts.end.start, fmts.end.end));

        appendComponentSection(
            parts, template, section(BRT_BEGIN_FONTS, BRT_END_FONTS), BRT_FONT, this.fonts,
        );
        appendComponentSection(
            parts, template, section(BRT_BEGIN_FILLS, BRT_END_FILLS), BRT_FILL, this.fills,
        );
        appendComponentSection(
            parts, template, section(BRT_BEGIN_BORDERS, BRT_END_BORDERS), BRT_BORDER, this.borders,
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
}

function appendComponentSection(
    parts: Uint8Array[],
    template: Uint8Array,
    section: { begin: Biff12RecordSpan; end: Biff12RecordSpan; inner: Biff12RecordSpan[] },
    componentId: number,
    custom: Uint8Array[],
): void {
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

/** Base XF index for a style that exactly matches one of the six base XFs. */
function baseXfIndex(style: CellStyle): number | null {
    const keys = Object.keys(style);
    if (keys.length === 1 && keys[0] === 'font'
        && style.font?.bold === true && Object.keys(style.font).length === 1) {
        return 3;
    }
    if (keys.length === 1 && keys[0] === 'fill'
        && style.fill?.patternType === 'gray125' && Object.keys(style.fill).length === 1) {
        return 4;
    }
    if (keys.length === 2 && style.font?.bold === true
        && Object.keys(style.font).length === 1
        && style.fill?.patternType === 'gray125' && Object.keys(style.fill).length === 1) {
        return 5;
    }
    return null;
}

// ---------------------------------------------------------------------------
// Reading styles.bin
// ---------------------------------------------------------------------------

const BUILTIN_NUMFMTS: Record<number, string> = {
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    14: 'mm-dd-yy',
    22: 'm/d/yy h:mm',
    49: '@',
};

function byteName(map: Record<string, number>, value: number): string | undefined {
    for (const [name, id] of Object.entries(map)) {
        if (id === value) return name === 'none' ? undefined : name;
    }
    return undefined;
}

function readColor(data: Uint8Array, offset: number): StyleColor | undefined {
    if (offset + 8 > data.length) return undefined;
    const kind = data[offset] >> 1;
    if (kind === 0) return undefined;
    if (kind === 1) return { indexed: data[offset + 1] };
    if (kind === 2) {
        const hex = (value: number): string => value.toString(16).padStart(2, '0').toUpperCase();
        return `${hex(data[offset + 7])}${hex(data[offset + 4])}${hex(data[offset + 5])}${hex(data[offset + 6])}`;
    }
    const view = new DataView(data.buffer, data.byteOffset + offset, 8);
    const tint = view.getInt16(2, true);
    const color: { theme: number; tint?: number } = { theme: data[offset + 1] };
    if (tint !== 0) color.tint = tint / 32767;
    return color;
}

function parseFontRecord(data: Uint8Array): FontStyle | undefined {
    if (data.length < 25) return undefined;
    const view = new DataView(data.buffer, data.byteOffset, data.length);
    const height = view.getUint16(0, true);
    const flags = view.getUint16(2, true);
    const weight = view.getUint16(4, true);
    const baseline = view.getUint16(6, true);
    const underline = data[8];
    const family = data[9];
    const charset = data[10];
    const color = readColor(data, 12);
    const scheme = data[20];
    const cch = view.getUint32(21, true);
    if (25 + cch * 2 > data.length) return undefined;

    const font: FontStyle = {};
    font.name = new TextDecoder('utf-16le').decode(data.subarray(25, 25 + cch * 2));
    font.size = height / 20;
    if ((flags & 0x01) !== 0 || weight >= 700) font.bold = true;
    if ((flags & 0x02) !== 0) font.italic = true;
    if ((flags & 0x08) !== 0) font.strike = true;
    if ((flags & 0x10) !== 0) font.outline = true;
    if ((flags & 0x20) !== 0) font.shadow = true;
    if ((flags & 0x40) !== 0) font.condense = true;
    if ((flags & 0x80) !== 0) font.extend = true;
    if (underline !== 0) {
        font.underline = underline === 0x21 ? 'singleAccounting'
            : underline === 0x22 ? 'doubleAccounting'
            : underline === 0x02 ? 'double'
            : 'single';
    }
    if (color !== undefined) font.color = color;
    if (family !== 0) font.family = family;
    if (charset !== 0) font.charset = charset;
    if (scheme === 1) font.scheme = 'major';
    else if (scheme === 2) font.scheme = 'minor';
    if (baseline === 1) font.vertAlign = 'superscript';
    else if (baseline === 2) font.vertAlign = 'subscript';
    return font;
}

function parseFillRecord(data: Uint8Array): FillStyle | undefined {
    if (data.length < 20) return undefined;
    const view = new DataView(data.buffer, data.byteOffset, data.length);
    const pattern = view.getUint32(0, true);
    const patternType = byteName(PATTERN_BYTES, pattern) as FillStyle['patternType'];
    const fgColor = readColor(data, 4);
    const bgColor = readColor(data, 12);
    if (patternType === undefined && fgColor === undefined && bgColor === undefined) return undefined;
    const fill: FillStyle = {};
    if (patternType !== undefined) fill.patternType = patternType;
    if (fgColor !== undefined) fill.fgColor = fgColor;
    if (bgColor !== undefined) fill.bgColor = bgColor;
    return fill;
}

function parseBorderRecord(data: Uint8Array): BorderStyle | undefined {
    if (data.length < 51) return undefined;
    const border: BorderStyle = {};
    if ((data[0] & 0x01) !== 0) border.diagonalDown = true;
    if ((data[0] & 0x02) !== 0) border.diagonalUp = true;
    const names = ['top', 'bottom', 'left', 'right', 'diagonal'] as const;
    for (let i = 0; i < names.length; i++) {
        const offset = 1 + i * 10;
        const style = byteName(BORDER_BYTES, data[offset]) as BorderSide['style'];
        const color = readColor(data, offset + 2);
        if (style !== undefined || color !== undefined) {
            const side: BorderSide = {};
            if (style !== undefined) side.style = style;
            if (color !== undefined) side.color = color;
            border[names[i]] = side;
        }
    }
    return Object.keys(border).length > 0 ? border : undefined;
}

/** Parse `xl/styles.bin` into per-cell-XF styles. */
export function parseXlsbStylesBin(data: Uint8Array): Array<CellStyle | null> {
    const records = scanRecords(data);
    const section = (beginId: number, endId: number): Biff12RecordSpan[] => {
        const begin = records.findIndex(record => record.id === beginId);
        const end = records.findIndex(record => record.id === endId);
        if (begin === -1 || end === -1 || end <= begin) return [];
        return records.slice(begin + 1, end);
    };

    const numFmts = new Map<number, string>();
    for (const record of section(BRT_BEGIN_FMTS, BRT_END_FMTS)) {
        if (record.id !== BRT_FMT || record.dataStart + 6 > record.end) continue;
        const view = new DataView(data.buffer, data.byteOffset + record.dataStart, record.end - record.dataStart);
        const id = view.getUint16(0, true);
        const cch = view.getUint32(2, true);
        if (6 + cch * 2 <= record.end - record.dataStart) {
            numFmts.set(id, new TextDecoder('utf-16le').decode(
                data.subarray(record.dataStart + 6, record.dataStart + 6 + cch * 2),
            ));
        }
    }

    const fonts: Array<FontStyle | undefined> = [];
    for (const record of section(BRT_BEGIN_FONTS, BRT_END_FONTS)) {
        if (record.id === BRT_FONT) {
            fonts.push(parseFontRecord(data.subarray(record.dataStart, record.end)));
        }
    }
    const fills: Array<FillStyle | undefined> = [];
    for (const record of section(BRT_BEGIN_FILLS, BRT_END_FILLS)) {
        if (record.id === BRT_FILL) {
            fills.push(parseFillRecord(data.subarray(record.dataStart, record.end)));
        }
    }
    const borders: Array<BorderStyle | undefined> = [];
    for (const record of section(BRT_BEGIN_BORDERS, BRT_END_BORDERS)) {
        if (record.id === BRT_BORDER) {
            borders.push(parseBorderRecord(data.subarray(record.dataStart, record.end)));
        }
    }

    const styles: Array<CellStyle | null> = [];
    for (const record of section(BRT_BEGIN_CELL_XFS, BRT_END_CELL_XFS)) {
        if (record.id !== BRT_XF || record.end - record.dataStart < 16) continue;
        const view = new DataView(data.buffer, data.byteOffset + record.dataStart, 16);
        const ifmt = view.getUint16(2, true);
        const fontId = view.getUint16(4, true);
        const fillId = view.getUint16(6, true);
        const borderId = view.getUint16(8, true);
        const rotation = data[record.dataStart + 10];
        const indent = data[record.dataStart + 11];
        const alignmentByte = data[record.dataStart + 12];
        const flagsByte = data[record.dataStart + 13];

        const style: CellStyle = {};
        if (ifmt !== 0) {
            const code = numFmts.get(ifmt) ?? BUILTIN_NUMFMTS[ifmt];
            if (code !== undefined) style.numFmt = code;
        }
        if (fontId !== 0 && fonts[fontId]) {
            style.font = fonts[fontId];
        }
        if (fillId > 1 && fills[fillId]) style.fill = fills[fillId];
        if (borderId !== 0 && borders[borderId]) style.border = borders[borderId];

        const horizontal = byteName(HORIZONTAL_BYTES, alignmentByte & 0x07);
        const vertical = byteName(VERTICAL_BYTES, (alignmentByte >> 3) & 0x07);
        const alignment: NonNullable<CellStyle['alignment']> = {};
        if (horizontal !== undefined && horizontal !== 'general') {
            alignment.horizontal = horizontal as NonNullable<CellStyle['alignment']>['horizontal'];
        }
        if (vertical !== undefined && vertical !== 'bottom') {
            alignment.vertical = vertical as NonNullable<CellStyle['alignment']>['vertical'];
        }
        if ((alignmentByte & 0x40) !== 0) alignment.wrapText = true;
        if ((alignmentByte & 0x80) !== 0) alignment.justifyLastLine = true;
        if ((flagsByte & 0x01) !== 0) alignment.shrinkToFit = true;
        const readingOrder = (flagsByte >> 2) & 0x03;
        if (readingOrder !== 0) alignment.readingOrder = readingOrder as 1 | 2;
        if (rotation !== 0) alignment.textRotation = rotation;
        if (indent !== 0) alignment.indent = indent;
        if (Object.keys(alignment).length > 0) style.alignment = alignment;

        const locked = (flagsByte & 0x10) !== 0;
        const hidden = (flagsByte & 0x20) !== 0;
        if (!locked || hidden) {
            const protection: NonNullable<CellStyle['protection']> = {};
            if (!locked) protection.locked = false;
            if (hidden) protection.hidden = true;
            style.protection = protection;
        }

        styles.push(Object.keys(style).length > 0 ? style : null);
    }

    return styles;
}
