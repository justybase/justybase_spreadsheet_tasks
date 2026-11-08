import {
    CellStyle,
    CellValue,
    ExcelError,
    FontStyle,
    PrimitiveCellValue,
    RichText,
    getCellStyle,
    isRichText,
    unwrapCell,
} from './Formats';

const OA_EPOCH_1900 = Date.UTC(1899, 11, 30);
const MARCH_1_1900 = Date.UTC(1900, 2, 1);

/**
 * Convert an Excel serial number to milliseconds since the given epoch.
 *
 * The 1900 date system counts a fictitious 1900-02-29 (serial 60), so serials
 * 1-59 denote 1900-01-01 through 1900-02-28 and must be shifted by one day
 * (Excel/xlspy/openpyxl parity). The 1904 system has no such gap.
 */
export function oaDateToTime(oaDate: number, epochMs: number): number {
    let serial = oaDate;
    if (epochMs === OA_EPOCH_1900 && serial >= 1 && serial < 60) {
        serial += 1;
    }
    return Math.round(serial * 86400000) + epochMs;
}

/**
 * Convert milliseconds since the given epoch to an Excel serial number.
 *
 * Inverse of {@link oaDateToTime}: dates before 1900-03-01 in the 1900 system
 * serialize one day lower to leave room for the fictitious 1900-02-29
 * (Excel/xlspy/openpyxl parity).
 */
export function timeToOaDate(timeMs: number, epochMs: number): number {
    let serial = (timeMs - epochMs) / 86400000;
    if (epochMs === OA_EPOCH_1900 && timeMs < MARCH_1_1900) {
        serial -= 1;
    }
    return serial;
}

/**
 * Whether a value carries a time of day. The writers store dates as the UTC
 * wall clock, so the UTC fields decide whether a datetime format is needed.
 */
export function dateHasTime(value: Date): boolean {
    return value.getUTCHours() !== 0 || value.getUTCMinutes() !== 0
        || value.getUTCSeconds() !== 0 || value.getUTCMilliseconds() !== 0;
}

/** Excel sheet name characters that must be replaced. */
export const INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;

/**
 * Sanitize an Excel sheet name (invalid chars, max 31 chars, non-empty).
 * @param sheetCount Current sheet count before adding this sheet (used for defaults).
 */
export function sanitizeSheetName(name: string, sheetCount: number): string {
    if (!name || typeof name !== 'string') {
        return `Sheet${sheetCount + 1}`;
    }

    let sanitized = name.replace(INVALID_SHEET_NAME_CHARS, '_');

    if (sanitized.length > 31) {
        sanitized = sanitized.substring(0, 31);
    }

    if (sanitized.trim().length === 0) {
        sanitized = `Sheet${sheetCount + 1}`;
    }

    return sanitized;
}

/** Excel limits (decision: throw on violation, like xlspy). */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;
export const MAX_CELL_CHARS = 32_767;

/** Length in UTF-16 code units (JS string length already counts astral as 2). */
export function utf16Length(text: string): number {
    return text.length;
}

/** Throw when cell text exceeds Excel's 32,767 UTF-16 unit limit. */
export function assertCellTextLength(text: string): void {
    if (text.length > MAX_CELL_CHARS) {
        throw new Error(
            `Excel cell text exceeds ${MAX_CELL_CHARS} UTF-16 code units (got ${text.length}).`
        );
    }
}

/** Throw when column count is outside Excel bounds. */
export function assertColumnCount(columnCount: number): void {
    if (!Number.isInteger(columnCount) || columnCount < 1 || columnCount > MAX_COLS) {
        throw new Error(
            `Excel column count must be 1..${MAX_COLS} (got ${columnCount}).`
        );
    }
}

/** Throw when a row/header is longer than the declared column count. */
export function assertRowFitsColumns(kind: string, length: number, columnCount: number): void {
    if (length > columnCount) {
        throw new Error(
            `Excel ${kind} length ${length} exceeds columnCount ${columnCount}.`
        );
    }
}

/** Throw when total row count exceeds Excel bounds. */
export function assertRowCount(totalRows: number): void {
    if (totalRows > MAX_ROWS) {
        throw new Error(
            `Excel row count exceeds ${MAX_ROWS} (got ${totalRows}).`
        );
    }
}

/** Create column-width scratch array filled with -1 (unknown). */
export function initColWidths(columnCount: number): number[] {
    return new Array(columnCount).fill(-1.0);
}

/** Excel's default column width in characters (Calibri 11, 96 DPI). */
export const DEFAULT_COLUMN_WIDTH = 8.43;
/** Hard cap for automatic column widths, in Excel character units. */
export const MAX_AUTOFIT_WIDTH = 200;
/**
 * Excel stores the column width plus the 5-pixel cell padding, i.e. 5/7 of a
 * character in Calibri 11 at 96 DPI. Automatic widths include it; explicit
 * `columnWidths` are written unchanged.
 */
export const COLUMN_WIDTH_PADDING = 5 / 7;
/** Room Excel reserves for the autofilter drop-down button, in characters. */
export const AUTOFILTER_BUTTON_WIDTH = 2.4;

/**
 * Rows measured by automatic widths. xlspy defaults to 1000; 200 keeps the
 * width estimate representative while cutting the write-time cost and the
 * streaming buffer (measured widths are dominated by the header and the first
 * rows anyway).
 */
export const AUTOFIT_SAMPLE_ROWS = 200;

/** Validate the `autofitRows` option (undefined: 200). */
export function resolveAutofitRows(value?: number): number {
    if (value === undefined) return AUTOFIT_SAMPLE_ROWS;
    if (!Number.isInteger(value) || value < 0) {
        throw new RangeError(`autofitRows must be a non-negative integer, got ${value}`);
    }
    return value;
}

/**
 * Safety factor over the widest advance xlspy measures for the fonts it
 * knows. Together with the per-character costs below it keeps automatic
 * widths at or above xlspy's, at the cost of some extra width.
 */
const AUTOFIT_FACTOR = 1.2;
/**
 * Excel gives long text 1/28 more room than short text (xlspy's
 * `_column_pixels`); apply it to the text part so long runs stay above xlspy.
 */
const LONG_TEXT_FACTOR = 29 / 28;
const BOLD_WIDTH_FACTOR = 1.08;
const ITALIC_WIDTH_FACTOR = 1.05;
const DEFAULT_FONT_SIZE = 11;
/** One level of Excel's alignment indent, in characters (9 px at 96 DPI). */
const INDENT_WIDTH = 9 / 7;

type WidthCosts = Record<string, number>;

function makeWidthCosts(groups: ReadonlyArray<readonly [string, number]>): WidthCosts {
    const costs: WidthCosts = Object.create(null);
    for (const [characters, cost] of groups) {
        for (const character of characters) costs[character] = cost;
    }
    return costs;
}

/**
 * Upper bounds of a character's advance in Calibri-11 digit widths. The
 * "narrow" table covers Calibri/Aptos/Arial Narrow (the default font family);
 * the "wide" table covers every other font xlspy knows, including unknown
 * fonts, so the estimate stays above xlspy for any styled font.
 */
const NARROW_FONT_COSTS = makeWidthCosts([
    ['@W', 1.75],
    ['%M', 1.7],
    ['m', 1.7],
    ['OQwGHN', 1.45],
    ['CD', 1.4],
    ['&U', 1.35],
    ['R', 1.25],
    ['BKAPT', 1.2],
    ['ESXbdpqu', 1.15],
    ['$0123456789Yhno', 1.1],
    ['#+<=>^`ae~', 1.1],
    ['?FZcg', 1.05],
    ['*_kLsvy|', 1.0],
    ['xz', 0.95],
    ['J', 0.9],
    ['"/', 0.85],
    ['\\r', 0.8],
    ['!-t', 0.75],
    ['()[]f{}', 0.7],
    ['I', 0.65],
    [',.:;', 0.65],
    ['l', 0.6],
    [' ij', 0.6],
    ["'", 0.5],
]);

const WIDE_FONT_COSTS = makeWidthCosts([
    ['@W', 2.5],
    ['M', 2.35],
    ['%', 2.3],
    ['m', 2.25],
    ['w', 2.05],
    ['&', 1.95],
    ['OQ', 1.9],
    ['NU', 1.85],
    ['GHD', 1.75],
    ['#K', 1.7],
    ['A', 1.7],
    ['+<=>^~', 1.65],
    ['V', 1.65],
    ['C', 1.6],
    ['BRXY', 1.6],
    ['Pbdhnopqu', 1.5],
    ['$0123456789e', 1.45],
    ['ESTZ', 1.45],
    ['*_ag', 1.4],
    ['FJLckvx|?`y{}', 1.35],
    ['"/\\', 1.3],
    ['s', 1.25],
    [" !'(),-.:;I[]fijlrtz", 1.25],
]);

const NARROW_FONT_NAMES = new Set([
    'calibri',
    'calibri light',
    'aptos',
    'aptos narrow',
    'arial narrow',
]);

function usesNarrowFont(name?: string | null): boolean {
    if (!name) return true;
    return NARROW_FONT_NAMES.has(name.trim().toLowerCase());
}

function isCjkWideCodePoint(cp: number): boolean {
    return (
        (cp >= 0x1100 && cp <= 0x115f) ||   // Hangul Jamo
        (cp >= 0x2e80 && cp <= 0x303e) ||   // CJK radicals and punctuation
        (cp >= 0x3041 && cp <= 0x33ff) ||   // Hiragana, Katakana, CJK compatibility
        (cp >= 0x3400 && cp <= 0x4dbf) ||   // CJK extension A
        (cp >= 0x4e00 && cp <= 0x9fff) ||   // CJK unified ideographs
        (cp >= 0xa000 && cp <= 0xa4cf) ||   // Yi
        (cp >= 0xac00 && cp <= 0xd7a3) ||   // Hangul syllables
        (cp >= 0xf900 && cp <= 0xfaff) ||   // CJK compatibility ideographs
        (cp >= 0xfe10 && cp <= 0xfe19) ||   // vertical forms
        (cp >= 0xfe30 && cp <= 0xfe6f) ||   // CJK compatibility forms
        (cp >= 0xff00 && cp <= 0xff60) ||   // fullwidth forms
        (cp >= 0xffe0 && cp <= 0xffe6) ||   // fullwidth signs
        (cp >= 0x20000 && cp <= 0x3fffd)    // CJK extensions B+
    );
}

function isEmojiCodePoint(cp: number): boolean {
    return (cp >= 0x1f000 && cp <= 0x1faff)  // emoji and pictographs
        || (cp >= 0x2300 && cp <= 0x23ff)   // technical symbols used as emoji
        || (cp >= 0x2600 && cp <= 0x27bf)   // miscellaneous symbols and dingbats
        || (cp >= 0x2b00 && cp <= 0x2bff);  // arrows and geometric shapes
}

function isZeroWidthCodePoint(cp: number): boolean {
    return cp < 0x20 || (cp >= 0x7f && cp <= 0x9f)
        || (cp >= 0x0300 && cp <= 0x036f)   // combining diacritics
        || (cp >= 0x1ab0 && cp <= 0x1aff)
        || (cp >= 0x1dc0 && cp <= 0x1dff)
        || (cp >= 0x20d0 && cp <= 0x20ff)
        || (cp >= 0xfe20 && cp <= 0xfe2f);
}

/**
 * Extra per-character costs for characters outside the Latin-1/ASCII tables,
 * generated from xlspy's widest font advances. Most non-ASCII letters are
 * covered by the range defaults below; this map only holds the outliers (wide
 * Latin Extended, Greek/Cyrillic and symbol characters).
 */
const WIDE_UNICODE_COSTS: WidthCosts = makeWidthCosts([
    ["\u00a5\u00c8\u00c9\u00ca\u00cb\u00de\u00e8\u00e9\u00ea\u00eb\u00f0\u00f1\u00f2\u00f3\u00f4\u00f5\u00f6\u00f8\u00f9\u00fa\u00fb\u00fc\u00fe\u0112\u0113\u0114\u0115\u0116\u0117\u0118\u0119\u011a\u011b\u0125\u0133\u0144\u0146\u0148\u014b\u014d\u014f\u0151\u015a\u015c\u015e\u0160\u0162\u0164\u0166\u0169\u016b\u016d\u016f\u0171\u0173\u0179\u017b\u017d\u0183\u018c\u018d\u018e\u019e\u01a5\u01a7\u01b4\u01bb\u01bc\u01bf\u01c9\u01d2\u01d4\u01d6\u01d8\u01da\u01dc\u01dd\u01eb\u01ed\u01f9\u01ff\u0204\u0205\u0206\u0207\u020d\u020f\u0215\u0217\u0218\u021a\u021f\u0222\u0223\u0228\u0229\u022b\u022d\u022f\u0231\u0232\u023b\u023e\u0246\u024b", 1.35],
    ["\u0127\u0185\u01ac\u01e8", 1.40],
    ["\u00a3\u00ab\u00b5\u00bb\u00c7\u00dd\u00df\u0106\u0108\u010a\u010c\u0111\u013f\u0154\u0156\u0158\u0176\u0178\u0180\u0182\u0184\u0186\u018b\u0198\u01a6\u01b3\u01e5\u0210\u0212\u0243\u024c", 1.45],
    ["\u00ac\u00b1\u00d7\u00f7\u0194\u01c2\u0245\u024e", 1.50],
    ["\u00b6\u00c0\u00c1\u00c2\u00c3\u00c4\u00c5\u0100\u0102\u0104\u010f\u0136\u0187\u01cd\u01de\u01e0\u01f7\u01fa\u0200\u0202\u0221\u0226\u023a", 1.55],
    ["\u00d0\u010e\u0110\u011c\u011e\u0120\u0122\u0124\u0149\u014a\u0189\u01a1\u01a4\u01e6\u01f4\u021e\u0220\u0235\u024a", 1.60],
    ["\u00d1\u00d2\u00d3\u00d4\u00d5\u00d6\u00d8\u00d9\u00da\u00db\u00dc\u0126\u0143\u0145\u0147\u014c\u014e\u0150\u0168\u016a\u016c\u016e\u0170\u0172\u0181\u018f\u0193\u019d\u019f\u01b0\u01b1\u01b2\u01c8\u01d1\u01d3\u01d5\u01d7\u01d9\u01db\u01e4\u01ea\u01ec\u01f8\u01fe\u020c\u020e\u0214\u0216\u022a\u022c\u022e\u0230\u0244", 1.70],
    ["\u00ae\u01a3", 1.80],
    ["\u018a", 1.85],
    ["\u0175\u01a0\u01af", 1.90],
    ["\u01e2", 1.95],
    ["\u00bc\u00bd\u00be\u0195\u01a2\u01c7\u01cc\u01e3", 2.00],
    ["\u0132\u03e3\u0416\u04c1\u04dc\u0502\u0503\u050b\u0906\u0911\u0912\u0913\u0914\u0974\u0975\u0e33\u1e3f\u1e41\u1e43\u20a5\u2106\u210b\u21ff\u220f\u226a\u226b", 2.05],
    ["\u00c6\u00e6\u01fc\u01fd\u0238\u0239\u03d6\u046d\u0518\u2105\u2154\u25ef", 2.10],
    ["\u00a9\u205b\u2117", 2.15],
    ["\u0152\u0153\u039c\u041c\u1e3e\u1e40\u1e42\u2180\u2182\u21ad\u21ae\u21ce\u21d4\u21dc\u21dd\u221e", 2.20],
    ["\u019c\u03e2\u0409\u040a\u0428\u042e\u0468\u0496\u04cd\u0514\u20a8\u2120\u2188", 2.25],
    ["\u0174\u01cb\u01f6\u0429\u04a6\u0508\u050a\u051c\u0520\u0522\u1e80\u1e82\u1e84\u1e86\u1e88\u2014\u2015\u2162\u2177\u2190\u2191\u2192\u2193\u2196\u2197\u2198\u2199\u222d", 2.30],
    ["\u01c6\u01f3\u046c\u0479\u047c\u20a9\u2133\u214f", 2.35],
    ["\u2230", 2.40],
    ["\u01ca\u2152", 2.45],
    ["\u0460\u047e", 2.50],
    ["\u01c5\u01f2\u052b\u2001\u2003\u25fd\u25fe", 2.55],
    ["\u0478\u2116\u2121\u213b\u22d8\u22d9", 2.65],
    ["\u2166\u216b", 2.70],
    ["\u01c4\u01f1\u20af", 2.75],
    ["\u2030", 3.00],
    ["\u052a\u20a7", 3.10],
    ["\u2194", 3.25],
    ["\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469\u246a\u246b\u246c\u246d\u246e\u246f\u2470\u2471\u2472\u2473\u24ea\u24eb\u24ec\u24ed\u24ee\u24ef\u24f0\u24f1\u24f2\u24f3\u24f4\u24ff", 3.35],
    ["\u2167", 3.45],
    ["\u2031", 3.60],
]);

/**
 * Conservative displayed-width units for `text`, calibrated so that the
 * resulting autofit width is not below xlspy's for the same text and font
 * (up to the 200-character cap). CJK/fullwidth characters count as 2.2 units,
 * emoji as 3.4, combining marks and control characters as zero. The optional
 * `fontName` selects the narrow (Calibri-like) or worst-case cost table.
 */
export function textDisplayUnits(text: string, fontName?: string | null): number {
    const costs = usesNarrowFont(fontName) ? NARROW_FONT_COSTS : WIDE_FONT_COSTS;
    let units = 0;
    for (const char of text) {
        const cp = char.codePointAt(0)!;
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
        if (cost !== undefined) {
            units += cost;
            continue;
        }
        // xlspy measures e.g. Hebrew, Arabic and Indic fallback glyphs at up
        // to ~2.0 digit widths; 1.6 was too optimistic.
        units += cp < 0x250 ? 1.3 : 2.0;
    }
    return units;
}

function fontScale(font?: FontStyle | null): number {
    let scale = Math.max(1, (font?.size ?? DEFAULT_FONT_SIZE) / DEFAULT_FONT_SIZE);
    if (font?.bold) scale *= BOLD_WIDTH_FACTOR;
    if (font?.italic) scale *= ITALIC_WIDTH_FACTOR;
    return scale;
}

/**
 * Characters Excel reserves for an alignment indent. One level is 9 px at
 * any font size; centred/distributed text is indented from both sides.
 */
function alignmentIndentChars(alignment?: CellStyle['alignment'] | null): number {
    const indent = alignment?.indent ?? 0;
    if (!Number.isFinite(indent) || indent <= 0) return 0;
    const sides = alignment?.horizontal === 'center' || alignment?.horizontal === 'distributed'
        ? 2
        : 1;
    return indent * sides * INDENT_WIDTH;
}

function autofitUnitsWidth(units: number, scale: number = 1, extraChars: number = 0): number {
    const width = (AUTOFIT_FACTOR * (units * LONG_TEXT_FACTOR + 1) + 2) * scale + extraChars;
    return Math.min(MAX_AUTOFIT_WIDTH, Math.max(DEFAULT_COLUMN_WIDTH, width));
}

/** Conservative autofit width in Excel character units (before file padding). */
export function autofitWidth(
    text: string,
    font?: FontStyle | null,
    indentChars: number = 0,
): number {
    return autofitUnitsWidth(textDisplayUnits(text, font?.name), fontScale(font), indentChars);
}

/** Autofit width as stored in the file: displayed width plus cell padding. */
export function autofitFileWidth(
    text: string,
    font?: FontStyle | null,
    indentChars: number = 0,
): number {
    // The 200-character cap is a cap on the value written to the file, so the
    // 5/7 padding must not push the stored width above it.
    return Math.min(
        MAX_AUTOFIT_WIDTH,
        autofitWidth(text, font, indentChars) + COLUMN_WIDTH_PADDING,
    );
}

function autofitFileUnitsWidth(units: number, indentChars: number = 0): number {
    return Math.min(
        MAX_AUTOFIT_WIDTH,
        autofitUnitsWidth(units, 1, indentChars) + COLUMN_WIDTH_PADDING,
    );
}

function richTextUnits(
    rich: RichText,
    baseFont?: FontStyle | null,
): number {
    let units = 0;
    for (const run of rich.runs) {
        const font = run.style?.font ?? baseFont;
        units += textDisplayUnits(run.text, font?.name) * fontScale(font);
    }
    return units;
}

/** Units of each line of a rich text, combining runs that cross lines. */
function richTextLineUnits(
    rich: RichText,
    baseFont?: FontStyle | null,
): number[] {
    const lines = [0];
    for (const run of rich.runs) {
        const font = run.style?.font ?? baseFont;
        const parts = run.text.split('\n');
        for (let index = 0; index < parts.length; index++) {
            if (index > 0) lines.push(0);
            lines[lines.length - 1] += textDisplayUnits(parts[index], font?.name) * fontScale(font);
        }
    }
    return lines;
}

const DATE_TOKEN_AT = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\.0+)/i;
const DATE_CHECK = /(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p)/i;
const DATE_SAMPLES: Record<string, string> = {
    yyyy: '2024',
    yy: '24',
    mmmmm: 'W',
    mmmm: 'październik',
    mmm: 'Sep',
    mm: '12',
    m: '12',
    dddd: 'poniedziałek',
    ddd: 'Wed',
    dd: '30',
    d: '30',
    hh: '23',
    h: '23',
    ss: '59',
    s: '59',
    'am/pm': 'PM',
    'a/p': 'P',
};

function formatSections(format: string): string[] {
    const sections: string[] = [];
    let section = '';
    let quoted = false;
    for (const char of format) {
        if (char === '"') quoted = !quoted;
        if (char === ';' && !quoted) {
            sections.push(section);
            section = '';
        } else {
            section += char;
        }
    }
    sections.push(section);
    return sections;
}

function withoutQuotedText(format: string): string {
    return format.replace(/"[^"]*"/g, '');
}

/** The text a date format section shows, using the longest locale samples xlspy uses. */
function dateFormatText(section: string): string | null {
    if (!DATE_CHECK.test(withoutQuotedText(section))) return null;
    let result = '';
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
        if (char === '\\' && i + 1 < section.length) {
            result += section[i + 1];
            i += 2;
            continue;
        }
        if (char === '[') {
            const end = section.indexOf(']', i + 1);
            i = end < 0 ? section.length : end + 1;
            continue;
        }
        const token = DATE_TOKEN_AT.exec(section.slice(i));
        if (token) {
            const key = token[1].toLowerCase();
            if (key.startsWith('.0')) {
                result += '.' + '0'.repeat(token[1].length - 1);
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

/** Literal characters a number format section displays besides the value. */
function formatLiteralText(section: string): string {
    let result = '';
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
        if (char === '\\' && i + 1 < section.length) {
            result += section[i + 1];
            i += 2;
            continue;
        }
        if (char === '[') {
            const end = section.indexOf(']', i + 1);
            if (end < 0) break;
            const content = section.slice(i + 1, end);
            if (content.startsWith('$')) {
                result += content.slice(1).replace(/-?\d+$/, '').trim();
            }
            i = end + 1;
            continue;
        }
        if ('0#?.,@'.includes(char)) {
            i++;
            continue;
        }
        if ('Ee+-'.includes(char)) {
            i++;
            continue;
        }
        result += char;
        i++;
    }
    return result;
}

function expandExponential(text: string): string {
    const match = /^(-?)(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/.exec(text);
    if (!match) return text;
    const sign = match[1];
    const integer = match[2];
    const fraction = match[3] ?? '';
    const digits = integer + fraction;
    const point = integer.length + Number(match[4]);
    if (point <= 0) return `${sign}0.${'0'.repeat(-point)}${digits}`;
    if (point >= digits.length) return `${sign}${digits}${'0'.repeat(point - digits.length)}`;
    return `${sign}${digits.slice(0, point)}.${digits.slice(point)}`;
}

/** Excel's General scientific form: up to 6 significant digits. */
function scientificText(value: number): string {
    const [mantissaRaw, exponentRaw] = Math.abs(value).toExponential(5).split('e');
    const mantissa = mantissaRaw.replace(/\.?0+$/, '');
    const exponent = Number(exponentRaw);
    return `${mantissa}E${exponent < 0 ? '-' : '+'}${String(Math.abs(exponent)).padStart(2, '0')}`;
}

/**
 * Excel's General display: at most 11 characters besides the sign, rounded
 * half-up, with a scientific fallback for numbers that do not fit or lose
 * significant digits. Mirrors xlspy's `_general_number` so long decimals and
 * large integers do not blow up the measured width.
 */
function generalNumberDisplay(value: number): string {
    if (!Number.isFinite(value)) {
        if (Number.isNaN(value)) return 'NaN';
        return value < 0 ? '-∞' : '∞';
    }
    if (value === 0) return '0';
    const sign = value < 0 ? '-' : '';
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
        if (fixed.includes('.')) fixed = fixed.replace(/0+$/, '').replace(/\.$/, '');
        const exact = rounded === abs;
        const room = whole > 0
            ? decimals + wholeLength
            : decimals + Math.floor(Math.log10(abs)) + 1;
        if (fixed !== '0' && (exact || room >= 6)) return sign + fixed;
    }
    return sign + scientificText(abs);
}

function generalNumberText(value: number, expandLarge: boolean): string {
    if (!Number.isFinite(value)) return String(value);
    let text = String(value);
    if (/[eE]/.test(text) && (expandLarge || Math.abs(value) < 1e-6)) {
        text = expandExponential(text);
    }
    if (text.endsWith('.0')) text = text.slice(0, -2);
    return text;
}

function generalValueText(value: PrimitiveCellValue, expandLarge: boolean = false): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'boolean') return 'PRAWDA';
    if (value instanceof Date) {
        return dateHasTime(value) ? '12/31/2024 23:45' : '12/31/2024';
    }
    if (typeof value === 'number') {
        return expandLarge ? generalNumberText(value, true) : generalNumberDisplay(value);
    }
    if (typeof value === 'bigint') return value.toString();
    if (value instanceof ExcelError) return value.code;
    return String(value);
}

/**
 * A conservative rendering of the value as Excel shows it, including date
 * names and number-format literals. It only needs to be wide enough for the
 * autofit estimate, so it intentionally over-counts rather than under-counts.
 */
function estimateCellDisplayText(value: PrimitiveCellValue, numFmt?: string | null): string {
    let raw = generalValueText(value);
    if (!numFmt || /^general$/i.test(numFmt.trim())) return raw;
    const sections = formatSections(numFmt);
    const negative = typeof value === 'bigint' ? value < 0n : typeof value === 'number' && value < 0;
    const secondSection = sections[1] ?? '';
    const useNegativeSection = negative && secondSection.trim() !== '';
    const section = useNegativeSection ? secondSection : (sections[0] ?? '');
    const displayValue: PrimitiveCellValue = useNegativeSection
        ? (typeof value === 'bigint' ? -value : Math.abs(value as number))
        : value;
    if (typeof displayValue !== 'number' && typeof displayValue !== 'bigint'
        && !(displayValue instanceof Date)) {
        return raw + formatLiteralText(section);
    }
    raw = generalValueText(displayValue, true);
    const dateText = dateFormatText(section);
    if (dateText !== null) {
        return textDisplayUnits(dateText) >= textDisplayUnits(raw) ? dateText : raw + dateText;
    }
    let result = raw + formatLiteralText(section);
    const clean = withoutQuotedText(section).replace(/\[[^\]]*\]/g, '');
    const placeholder = /[0#?][0#?,.]*/.exec(clean);
    if (placeholder) {
        const [integerPart, fractionPart = ''] = placeholder[0].split('.');
        const minIntegerDigits = (integerPart.match(/0/g) ?? []).length;
        const integerMatch = /^-?\d+/.exec(raw);
        const rawIntegerDigits = integerMatch ? integerMatch[0].replace('-', '').length : 0;
        result += '0'.repeat(Math.max(0, minIntegerDigits - rawIntegerDigits));
        const decimals = (fractionPart.match(/[0#?]/g) ?? []).length;
        const rawFractionDigits = raw.includes('.') ? raw.split('.')[1].length : 0;
        const extraDecimals = Math.max(0, decimals - rawFractionDigits);
        if (extraDecimals > 0) {
            result += (raw.includes('.') ? '' : '.') + '0'.repeat(extraDecimals);
        }
        if (integerPart.includes(',')) {
            const digits = (raw.match(/\d/g) ?? []).length;
            result += ','.repeat(Math.max(0, Math.floor((digits - 1) / 3)));
        }
    }
    const exponent = /[eE][+-]([0#]+)/.exec(clean);
    if (exponent) result += 'E+' + '0'.repeat(Math.max(2, exponent[1].length));
    const percent = (clean.match(/%/g) ?? []).length;
    if (percent > 0) result += '00'.repeat(percent) + '%'.repeat(percent);
    return result;
}

/** Options for {@link applyHeaderWidths}. */
export interface HeaderWidthOptions {
    /** Header cells use a bold font. */
    bold?: boolean;
    /** Add room for the autofilter drop-down button. */
    filterButton?: boolean;
}

/** Apply autofit widths from header labels. */
export function applyHeaderWidths(
    colWidths: number[],
    headers: string[],
    columnCount: number,
    options?: HeaderWidthOptions,
): void {
    const font = options?.bold ? { bold: true } : null;
    for (let i = 0; i < columnCount; i++) {
        let width = autofitFileWidth(headers[i] ?? '', font);
        if (options?.filterButton) {
            width = Math.min(MAX_AUTOFIT_WIDTH, width + AUTOFILTER_BUTTON_WIDTH);
        }
        if (colWidths[i] < width) colWidths[i] = width;
    }
}

function richTextFileWidth(
    rich: RichText,
    baseFont?: FontStyle | null,
    alignment?: CellStyle['alignment'] | null,
): number {
    const indent = alignmentIndentChars(alignment);
    if (alignment?.wrapText && rich.runs.some(run => run.text.includes('\n'))) {
        let best = 0;
        for (const units of richTextLineUnits(rich, baseFont)) {
            best = Math.max(best, autofitFileUnitsWidth(units, indent));
        }
        return best;
    }
    return autofitFileUnitsWidth(richTextUnits(rich, baseFont), indent);
}

/** Autofit file width of one cell, honouring its style and rich-text runs. */
function cellAutofitFileWidth(cell: CellValue): number {
    if (isRichText(cell)) return richTextFileWidth(cell);
    const style = getCellStyle(cell);
    const value = unwrapCell(cell);
    if (value === null || value === undefined) return -1;
    const indent = alignmentIndentChars(style?.alignment);
    if (isRichText(value)) return richTextFileWidth(value, style?.font, style?.alignment);
    const text = estimateCellDisplayText(value, style?.numFmt);
    if (style?.alignment?.wrapText && text.includes('\n')) {
        let best = 0;
        for (const line of text.split('\n')) {
            best = Math.max(best, autofitFileWidth(line, style?.font, indent));
        }
        return best;
    }
    return autofitFileWidth(text, style?.font, indent);
}

/**
 * Sample up to `sampleLimit` rows and expand column widths (autofit).
 * Skips null/undefined cells; understands styled, formula and rich-text cells.
 */
export function updateColWidthsFromRows(
    colWidths: number[],
    rows: unknown[][],
    sampleLimit: number = AUTOFIT_SAMPLE_ROWS
): void {
    for (let r = 0; r < Math.min(rows.length, sampleLimit); r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
            const raw = row[c];
            if (raw === null || raw === undefined) continue;
            const width = cellAutofitFileWidth(raw as CellValue);
            if (width >= 0 && colWidths[c] < width) colWidths[c] = width;
        }
    }
}

/**
 * Resolve a stored width to the value written into the file. The fallback is
 * Excel's default width plus cell padding, matching the old 10-character
 * default while staying close to xlspy/Excel.
 */
export function defaultColWidth(width: number, floor: boolean = false): number {
    const value = width > 0 ? width : DEFAULT_COLUMN_WIDTH + COLUMN_WIDTH_PADDING;
    return floor ? Math.floor(value) : value;
}

/** Validate one explicit column width (xlspy accepts 0.5..255 characters). */
export function assertColumnWidth(width: number, index: number): number {
    if (typeof width !== 'number' || !Number.isFinite(width)) {
        throw new TypeError(`columnWidths[${index}] must be a finite number`);
    }
    if (width < 0.5 || width > 255) {
        throw new RangeError(`columnWidths[${index}] must be between 0.5 and 255, got ${width}`);
    }
    return width;
}

/** Compact XML attribute value for an XLSX column width (2 decimals max). */
export function formatColumnWidth(width: number): string {
    return String(Math.round(width * 100) / 100);
}

/**
 * The `r` attribute of a cell (leading space included), or '' when the cell
 * directly follows the previous written cell in its row. Such cells take their
 * column from their position, so `r` is only needed after a gap. The first cell
 * of a row has `prevCol = -1` and sits in column A without `r`.
 */
export function cellRefAttr(col: number, prevCol: number, colRef: string, rowNum: number): string {
    return col === prevCol + 1 ? '' : ` r="${colRef}${rowNum}"`;
}

/** Deflate level for the ZIP package. 3 trades speed for size; 6 is zlib's default and made XLSB writes about 2.6x slower for an 8% smaller file. */
export const DEFAULT_COMPRESSION_LEVEL = 3;

export function assertCompressionLevel(level: number): number {
    if (!Number.isInteger(level) || level < 0 || level > 9) {
        throw new RangeError(`compressionLevel must be an integer from 0 to 9, got ${level}`);
    }
    return level;
}
