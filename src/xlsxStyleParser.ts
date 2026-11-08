/**
 * Read-side helpers for XLSX `styles.xml` and rich shared strings.
 *
 * The writer emits a known, simple subset of styles; this parser also accepts
 * the equivalent Excel markup (attribute order, quotes, self-closing tags).
 */
import {
    AlignmentStyle,
    BorderSide,
    BorderStyle,
    CellStyle,
    FillStyle,
    FontStyle,
    ProtectionStyle,
    RichText,
    StyleColor,
} from './Formats';
import { unescapeXmlText } from './xmlUtils';

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

function attribute(tag: string, name: string): string | undefined {
    const double = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`).exec(tag);
    if (double) return double[1];
    const single = new RegExp(`\\b${name}\\s*=\\s*'([^']*)'`).exec(tag);
    return single?.[1];
}

function parseColorTag(colorTag: string): StyleColor | undefined {
    if (!colorTag) return undefined;
    const rgb = attribute(colorTag, 'rgb');
    if (rgb !== undefined) return rgb;
    const indexed = attribute(colorTag, 'indexed');
    if (indexed !== undefined) return { indexed: parseInt(indexed, 10) };
    const theme = attribute(colorTag, 'theme');
    if (theme !== undefined) {
        const tint = attribute(colorTag, 'tint');
        const color: { theme: number; tint?: number } = { theme: parseInt(theme, 10) };
        if (tint !== undefined) color.tint = parseFloat(tint);
        return color;
    }
    return undefined;
}

function colorFrom(tag: string): StyleColor | undefined {
    const colorTag = /<color\b[^>]*\/?>/.exec(tag)?.[0];
    return parseColorTag(colorTag ?? '');
}

function flagValue(block: string, name: string): boolean {
    const match = new RegExp(`<${name}(?:\\s+val\\s*=\\s*["']([^"']*)["'])?\\s*/?>`).exec(block);
    if (!match) return false;
    const value = match[1];
    return value === undefined || !/^(?:0|false)$/i.test(value);
}

function numberAttribute(tag: string, name: string): number | undefined {
    const value = attribute(tag, name);
    if (value === undefined) return undefined;
    const parsed = parseFloat(value);
    return Number.isFinite(parsed) ? parsed : undefined;
}

/** Parse the inside of `<font>` or `<rPr>` into a font style. */
export function parseXlsxFont(block: string): FontStyle | undefined {
    const font: FontStyle = {};
    const name = /<name\b[^>]*\bval\s*=\s*["']([^"']*)["']/.exec(block)?.[1];
    if (name !== undefined) font.name = unescapeXmlText(name);
    const size = numberAttribute(/<sz\b[^>]*\/?>/.exec(block)?.[0] ?? '', 'val');
    if (size !== undefined) font.size = size;
    if (flagValue(block, 'b')) font.bold = true;
    if (flagValue(block, 'i')) font.italic = true;
    if (flagValue(block, 'strike')) font.strike = true;
    if (flagValue(block, 'outline')) font.outline = true;
    if (flagValue(block, 'shadow')) font.shadow = true;
    if (flagValue(block, 'condense')) font.condense = true;
    if (flagValue(block, 'extend')) font.extend = true;
    const underlineTag = /<u\b[^>]*\/?>/.exec(block)?.[0];
    if (underlineTag) {
        const value = attribute(underlineTag, 'val');
        if (value === undefined || value === 'single') font.underline = 'single';
        else if (value !== 'none') font.underline = value as FontStyle['underline'];
    }
    const color = colorFrom(block);
    if (color !== undefined) {
        font.color = color;
    }
    const family = numberAttribute(/<family\b[^>]*\/?>/.exec(block)?.[0] ?? '', 'val');
    if (family !== undefined) font.family = family;
    const charset = numberAttribute(/<charset\b[^>]*\/?>/.exec(block)?.[0] ?? '', 'val');
    if (charset !== undefined) font.charset = charset;
    const scheme = /<scheme\b[^>]*\bval\s*=\s*["']([^"']*)["']/.exec(block)?.[1];
    if (scheme === 'major' || scheme === 'minor') font.scheme = scheme;
    const vertAlign = /<vertAlign\b[^>]*\bval\s*=\s*["']([^"']*)["']/.exec(block)?.[1];
    if (vertAlign === 'superscript' || vertAlign === 'subscript') font.vertAlign = vertAlign;
    return Object.keys(font).length > 0 ? font : undefined;
}

function parseFill(block: string): FillStyle | undefined {
    const patternTag = /<patternFill\b[^>]*>/.exec(block)?.[0] ?? '';
    const patternType = attribute(patternTag, 'patternType') as FillStyle['patternType'];
    const fgColor = parseColorTag(/<fgColor\b[^>]*\/?>/.exec(block)?.[0] ?? '');
    const bgColor = parseColorTag(/<bgColor\b[^>]*\/?>/.exec(block)?.[0] ?? '');
    if (patternType === undefined && fgColor === undefined && bgColor === undefined) return undefined;
    const fill: FillStyle = {};
    if (patternType !== undefined) fill.patternType = patternType;
    if (fgColor !== undefined) fill.fgColor = fgColor;
    if (bgColor !== undefined) fill.bgColor = bgColor;
    return fill;
}

function parseBorderSide(block: string, tag: string): BorderSide | undefined {
    const match = new RegExp(`<${tag}(\\s[^>]*)?>([\\s\\S]*?)<\\/${tag}>|<${tag}(\\s[^>]*)?\\/>`).exec(block);
    if (!match) return undefined;
    const openTag = `<${tag}${match[1] ?? match[3] ?? ''}>`;
    const style = attribute(openTag, 'style') as BorderSide['style'];
    const color = colorFrom(match[2] ?? '');
    if (style === undefined && color === undefined) return undefined;
    const side: BorderSide = {};
    if (style !== undefined) side.style = style;
    if (color !== undefined) side.color = color;
    return side;
}

function parseBorder(block: string): BorderStyle | undefined {
    const openTag = /<border\b[^>]*>/.exec(block)?.[0] ?? '';
    const border: BorderStyle = {};
    for (const name of ['top', 'bottom', 'left', 'right', 'diagonal'] as const) {
        const side = parseBorderSide(block, name);
        if (side) border[name] = side;
    }
    if (/diagonalUp\s*=\s*["'](?:1|true)["']/i.test(openTag)) border.diagonalUp = true;
    if (/diagonalDown\s*=\s*["'](?:1|true)["']/i.test(openTag)) border.diagonalDown = true;
    if (/outline\s*=\s*["'](?:0|false)["']/i.test(openTag)) border.outline = false;
    return Object.keys(border).length > 0 ? border : undefined;
}

function parseAlignment(tag: string | undefined): AlignmentStyle | undefined {
    if (!tag) return undefined;
    const alignment: AlignmentStyle = {};
    const horizontal = attribute(tag, 'horizontal');
    if (horizontal !== undefined) alignment.horizontal = horizontal as AlignmentStyle['horizontal'];
    const vertical = attribute(tag, 'vertical');
    if (vertical !== undefined) alignment.vertical = vertical as AlignmentStyle['vertical'];
    if (/wrapText\s*=\s*["'](?:1|true)["']/i.test(tag)) alignment.wrapText = true;
    if (/shrinkToFit\s*=\s*["'](?:1|true)["']/i.test(tag)) alignment.shrinkToFit = true;
    const indent = numberAttribute(tag, 'indent');
    if (indent !== undefined) alignment.indent = indent;
    const rotation = numberAttribute(tag, 'textRotation');
    if (rotation !== undefined) alignment.textRotation = rotation;
    if (/justifyLastLine\s*=\s*["'](?:1|true)["']/i.test(tag)) alignment.justifyLastLine = true;
    const readingOrder = numberAttribute(tag, 'readingOrder');
    if (readingOrder !== undefined) alignment.readingOrder = readingOrder as 0 | 1 | 2;
    return Object.keys(alignment).length > 0 ? alignment : undefined;
}

function parseProtection(tag: string | undefined): ProtectionStyle | undefined {
    if (!tag) return undefined;
    const protection: ProtectionStyle = {};
    if (/locked\s*=\s*["'](?:0|false)["']/i.test(tag)) protection.locked = false;
    if (/hidden\s*=\s*["'](?:1|true)["']/i.test(tag)) protection.hidden = true;
    return Object.keys(protection).length > 0 ? protection : undefined;
}

export interface XlsxStyleTable {
    /** Cell style per `cellXfs` index. */
    cellStyles: Array<CellStyle | null>;
    /** Set of XF indices whose number format is a date. */
    dateXfs: Set<number>;
}

/** Parse `xl/styles.xml` into per-XF cell styles. */
export function parseXlsxStyles(xml: string): XlsxStyleTable {
    const numFmts = new Map<number, string>();
    const numFmtRegex = /<numFmt\b[^>]*>/g;
    let match: RegExpExecArray | null;
    while ((match = numFmtRegex.exec(xml)) !== null) {
        const id = numberAttribute(match[0], 'numFmtId');
        const code = attribute(match[0], 'formatCode');
        if (id !== undefined && code !== undefined) numFmts.set(id, unescapeXmlText(code));
    }

    const fonts: Array<FontStyle | undefined> = [];
    const fontsBlock = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(xml)?.[1] ?? '';
    const fontRegex = /<font\b[^>]*>([\s\S]*?)<\/font>/g;
    while ((match = fontRegex.exec(fontsBlock)) !== null) fonts.push(parseXlsxFont(match[1]));

    const fills: Array<FillStyle | undefined> = [];
    const fillsBlock = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(xml)?.[1] ?? '';
    const fillRegex = /<fill\b[^>]*>([\s\S]*?)<\/fill>/g;
    while ((match = fillRegex.exec(fillsBlock)) !== null) fills.push(parseFill(match[1]));

    const borders: Array<BorderStyle | undefined> = [];
    const bordersBlock = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(xml)?.[1] ?? '';
    const borderRegex = /<border\b[^>]*>([\s\S]*?)<\/border>/g;
    while ((match = borderRegex.exec(bordersBlock)) !== null) borders.push(parseBorder(match[0]));

    const cellStyles: Array<CellStyle | null> = [];
    const dateXfs = new Set<number>();
    const xfsBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? '';
    const xfRegex = /<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g;
    let xfIndex = 0;
    while ((match = xfRegex.exec(xfsBlock)) !== null) {
        const openTag = `<xf${match[1]}>`;
        const children = match[2] ?? '';
        const style: CellStyle = {};
        const numFmtId = numberAttribute(openTag, 'numFmtId') ?? 0;
        if (numFmtId !== 0) {
            const code = numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS[numFmtId];
            if (code !== undefined) style.numFmt = code;
        }
        const fontId = numberAttribute(openTag, 'fontId') ?? 0;
        if (fontId !== 0 && fonts[fontId]) {
            style.font = fonts[fontId];
        }
        const fillId = numberAttribute(openTag, 'fillId') ?? 0;
        if (fillId !== 0) {
            const fill = fills[fillId];
            if (fill && fill.patternType !== 'none' && fill.patternType !== 'gray125') {
                style.fill = fill;
            }
        }
        const borderId = numberAttribute(openTag, 'borderId') ?? 0;
        if (borderId !== 0 && borders[borderId]) style.border = borders[borderId];
        const alignment = parseAlignment(/<alignment\b[^>]*\/?>/.exec(children)?.[0]);
        const protection = parseProtection(/<protection\b[^>]*\/?>/.exec(children)?.[0]);
        if (alignment) style.alignment = alignment;
        if (protection) style.protection = protection;

        const isDate = numFmtId >= 14 && numFmtId <= 22;
        if (isDate) dateXfs.add(xfIndex);
        cellStyles.push(Object.keys(style).length > 0 ? style : null);
        xfIndex++;
    }

    return { cellStyles, dateXfs };
}

/** Parse `<r>` runs inside a shared-string `<si>` or an inline `<is>`. */
function parseRichRuns(content: string): Array<{ text: string; style?: CellStyle }> {
    const runs: Array<{ text: string; style?: CellStyle }> = [];
    const runRegex = /<r\b[^>]*>([\s\S]*?)<\/r>/g;
    let run: RegExpExecArray | null;
    while ((run = runRegex.exec(content)) !== null) {
        const fontBlock = /<rPr\b[^>]*>([\s\S]*?)<\/rPr>/.exec(run[1])?.[1];
        const text = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(run[1])?.[1] ?? '';
        const runStyle = fontBlock ? parseXlsxFont(fontBlock) : undefined;
        runs.push({
            text: unescapeXmlText(text),
            style: runStyle ? { font: runStyle } : undefined,
        });
    }
    return runs;
}

/** Parse rich-text runs from an inline-string `<is>` body. */
export function parseXlsxInlineRichText(content: string): RichText | null {
    const runs = parseRichRuns(content.replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, ''));
    return runs.length > 0 ? new RichText(runs) : null;
}

/** Parse rich-text runs from `xl/sharedStrings.xml`; plain entries are null. */
export function parseXlsxRichTexts(xml: string): Array<RichText | null> {
    const result: Array<RichText | null> = [];
    const itemRegex = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let item: RegExpExecArray | null;
    while ((item = itemRegex.exec(xml)) !== null) {
        const content = (item[1] ?? '')
            .replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, '')
            .replace(/<rPh\b[^>]*\/>/g, '');
        const runs = parseRichRuns(content);
        result.push(runs.length > 0 ? new RichText(runs) : null);
    }
    return result;
}
