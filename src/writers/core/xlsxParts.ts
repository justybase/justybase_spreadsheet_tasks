/**
 * Shared XLSX package serialization.
 *
 * Everything in this module is pure string/byte generation with no Node or
 * browser imports: the Node writer and the browser writer both call it, which
 * is what keeps their unpacked parts identical.
 */
import {
    AlignmentStyle,
    BorderSide,
    BorderStyle,
    CellStyle,
    CellValue,
    ExcelError,
    FillStyle,
    FontStyle,
    FormulaCell,
    ProtectionStyle,
    RichText,
    StyleColor,
    getCellStyle,
    getFormula,
    isRichText,
    unwrapCell,
} from '../../Formats';
import {
    cellRefAttr,
    dateHasTime,
    defaultColWidth,
    formatColumnWidth,
} from '../../writerHelpers';
import { escapeXmlText, escapeXlsxText } from '../../xmlUtils';
import { TextByteSink } from './ByteSink';

export interface XlsxSheetEntry {
    name: string;
    pathInArchive: string;
    nameInArchive: string;
    hidden: boolean;
    sheetId: number;
    rId: string;
    filterHeaderRange: string | null;
}

/** Cell styles 0-5 are always emitted; custom styles start at 6. */
export const XLSX_BASE_XF_COUNT = 6;

export function escapeXlsxAttribute(value: string): string {
    return escapeXmlText(value);
}

/** Sheet name as used inside a defined name / formula. */
export function formatSheetNameForFormula(name: string): string {
    if (/[\s\-+=()!@#$%^&]/.test(name) || /^[0-9]/.test(name)) {
        return `'${name.replace(/'/g, "''")}'`;
    }
    return name;
}

const BORDER_LINE_STYLES = new Set([
    'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair',
    'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot',
    'mediumDashDotDot', 'slantDashDot',
]);

const PATTERN_TYPES = new Set([
    'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal',
    'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
    'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid',
    'lightTrellis', 'gray125', 'gray0625',
]);

const HORIZONTAL_ALIGNMENTS = new Set([
    'general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed',
]);

const VERTICAL_ALIGNMENTS = new Set(['top', 'center', 'bottom', 'justify', 'distributed']);

function colorKey(color: StyleColor): string {
    if (typeof color === 'string') {
        let value = color.trim().replace(/^#/, '').toUpperCase();
        if (value.length === 6) value = `FF${value}`;
        if (!/^[0-9A-F]{8}$/.test(value)) {
            throw new Error(`Invalid color '${color}'. Use RRGGBB, AARRGGBB or '#' + hex.`);
        }
        return `rgb:${value}`;
    }
    if ('indexed' in color) return `indexed:${color.indexed}`;
    const tint = color.tint ?? 0;
    return `theme:${color.theme}:${tint}`;
}

function colorXml(color: StyleColor | undefined): string {
    if (color === undefined) return '';
    if (typeof color === 'string') {
        const key = colorKey(color);
        return `<color rgb="${key.slice(4)}"/>`;
    }
    if ('indexed' in color) return `<color indexed="${color.indexed}"/>`;
    const tint = color.tint;
    return `<color theme="${color.theme}"${tint !== undefined ? ` tint="${tint}"` : ''}/>`;
}

function normalizeFont(font: FontStyle | undefined): FontStyle | null {
    if (!font) return null;
    const out: FontStyle = {};
    if (font.name !== undefined) out.name = font.name;
    if (font.size !== undefined) {
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
    if (font.underline === true) out.underline = 'single';
    else if (typeof font.underline === 'string') out.underline = font.underline;
    if (font.color !== undefined) out.color = font.color;
    if (font.family !== undefined) out.family = font.family;
    if (font.scheme !== undefined) out.scheme = font.scheme;
    if (font.vertAlign !== undefined && font.vertAlign !== 'baseline') out.vertAlign = font.vertAlign;
    if (font.charset !== undefined) out.charset = font.charset;
    return Object.keys(out).length > 0 ? out : null;
}

function normalizeFill(fill: FillStyle | undefined): FillStyle | null {
    if (!fill) return null;
    const out: FillStyle = {};
    if (fill.patternType !== undefined) {
        if (!PATTERN_TYPES.has(fill.patternType)) throw new Error(`Unknown fill pattern '${fill.patternType}'`);
        out.patternType = fill.patternType;
    }
    if (fill.fgColor !== undefined) out.fgColor = fill.fgColor;
    if (fill.bgColor !== undefined) out.bgColor = fill.bgColor;
    if (Object.keys(out).length === 0) return null;
    if (out.patternType === undefined) out.patternType = 'solid';
    return out;
}

function normalizeBorderSide(side: BorderSide | undefined): BorderSide | null {
    if (!side) return null;
    if (side.style !== undefined && !BORDER_LINE_STYLES.has(side.style)) {
        throw new Error(`Unknown border style '${side.style}'`);
    }
    const out: BorderSide = {};
    if (side.style !== undefined) out.style = side.style;
    if (side.color !== undefined) out.color = side.color;
    return Object.keys(out).length > 0 ? out : null;
}

function normalizeBorder(border: BorderStyle | undefined): BorderStyle | null {
    if (!border) return null;
    const out: BorderStyle = {};
    for (const name of ['top', 'bottom', 'left', 'right', 'diagonal'] as const) {
        const side = normalizeBorderSide(border[name]);
        if (side) out[name] = side;
    }
    if (border.diagonalUp) out.diagonalUp = true;
    if (border.diagonalDown) out.diagonalDown = true;
    if (border.outline === false) out.outline = false;
    return Object.keys(out).length > 0 ? out : null;
}

function normalizeAlignment(alignment: AlignmentStyle | undefined): AlignmentStyle | null {
    if (!alignment) return null;
    const out: AlignmentStyle = {};
    if (alignment.horizontal !== undefined) {
        if (!HORIZONTAL_ALIGNMENTS.has(alignment.horizontal)) {
            throw new Error(`Unknown horizontal alignment '${alignment.horizontal}'`);
        }
        out.horizontal = alignment.horizontal;
    }
    if (alignment.vertical !== undefined) {
        if (!VERTICAL_ALIGNMENTS.has(alignment.vertical)) {
            throw new Error(`Unknown vertical alignment '${alignment.vertical}'`);
        }
        out.vertical = alignment.vertical;
    }
    if (alignment.wrapText) out.wrapText = true;
    if (alignment.shrinkToFit) out.shrinkToFit = true;
    if (alignment.indent !== undefined) {
        if (!Number.isInteger(alignment.indent) || alignment.indent < 0 || alignment.indent > 250) {
            throw new RangeError(`alignment.indent must be 0..250, got ${alignment.indent}`);
        }
        out.indent = alignment.indent;
    }
    if (alignment.textRotation !== undefined) {
        const rotation = alignment.textRotation;
        if (!Number.isInteger(rotation) || !((rotation >= 0 && rotation <= 180) || rotation === 255)) {
            throw new RangeError(`alignment.textRotation must be 0..180 or 255, got ${rotation}`);
        }
        out.textRotation = rotation;
    }
    if (alignment.justifyLastLine) out.justifyLastLine = true;
    if (alignment.readingOrder !== undefined) out.readingOrder = alignment.readingOrder;
    return Object.keys(out).length > 0 ? out : null;
}

function normalizeProtection(protection: ProtectionStyle | undefined): ProtectionStyle | null {
    if (!protection) return null;
    const out: ProtectionStyle = {};
    if (protection.locked === false) out.locked = false;
    if (protection.hidden) out.hidden = true;
    return Object.keys(out).length > 0 ? out : null;
}

/** Canonical form of a style; equal styles get identical JSON keys. */
export function normalizeStyle(style: CellStyle | null | undefined): CellStyle | null {
    if (!style) return null;
    const out: CellStyle = {};
    if (style.numFmt !== undefined && style.numFmt !== '') out.numFmt = style.numFmt;
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

const BASE_FONT_NORMAL =
    '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const BASE_FONT_BOLD =
    '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const BASE_FILL_NONE = '<fill><patternFill patternType="none"/></fill>';
const BASE_FILL_GRAY = '<fill><patternFill patternType="gray125"/></fill>';
const BASE_FILL_SOLID =
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>';
const BASE_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

const BASE_XFS = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
];

function fontXml(font: FontStyle): string {
    let xml = '<font>';
    if (font.name !== undefined) xml += `<name val="${escapeXlsxAttribute(font.name)}"/>`;
    if (font.charset !== undefined) xml += `<charset val="${font.charset}"/>`;
    if (font.family !== undefined) xml += `<family val="${font.family}"/>`;
    if (font.bold) xml += '<b/>';
    if (font.italic) xml += '<i/>';
    if (font.strike) xml += '<strike/>';
    if (font.outline) xml += '<outline/>';
    if (font.shadow) xml += '<shadow/>';
    if (font.condense) xml += '<condense/>';
    if (font.extend) xml += '<extend/>';
    if (font.color !== undefined) xml += colorXml(font.color);
    if (font.size !== undefined) xml += `<sz val="${font.size}"/>`;
    if (font.underline !== undefined) {
        xml += font.underline === 'single' ? '<u/>' : `<u val="${font.underline}"/>`;
    }
    if (font.vertAlign !== undefined) xml += `<vertAlign val="${font.vertAlign}"/>`;
    if (font.scheme !== undefined) xml += `<scheme val="${font.scheme}"/>`;
    return `${xml}</font>`;
}

function fillXml(fill: FillStyle): string {
    const pattern = fill.patternType ?? 'solid';
    let xml = `<fill><patternFill patternType="${pattern}">`;
    if (fill.fgColor !== undefined) xml += colorXml(fill.fgColor).replace('<color', '<fgColor');
    if (fill.bgColor !== undefined) xml += colorXml(fill.bgColor).replace('<color', '<bgColor');
    return `${xml}</patternFill></fill>`;
}

function borderSideXml(tag: string, side: BorderSide | undefined): string {
    if (!side || (side.style === undefined && side.color === undefined)) return `<${tag}/>`;
    let xml = `<${tag}`;
    if (side.style !== undefined) xml += ` style="${side.style}"`;
    xml += '>';
    if (side.color !== undefined) xml += colorXml(side.color);
    return `${xml}</${tag}>`;
}

function borderXml(border: BorderStyle): string {
    let xml = '<border';
    if (border.diagonalUp) xml += ' diagonalUp="1"';
    if (border.diagonalDown) xml += ' diagonalDown="1"';
    if (border.outline === false) xml += ' outline="0"';
    xml += '>';
    xml += borderSideXml('left', border.left);
    xml += borderSideXml('right', border.right);
    xml += borderSideXml('top', border.top);
    xml += borderSideXml('bottom', border.bottom);
    xml += borderSideXml('diagonal', border.diagonal);
    return `${xml}</border>`;
}

function alignmentXml(alignment: AlignmentStyle): string {
    let xml = '<alignment';
    if (alignment.horizontal !== undefined) xml += ` horizontal="${alignment.horizontal}"`;
    if (alignment.vertical !== undefined) xml += ` vertical="${alignment.vertical}"`;
    if (alignment.wrapText) xml += ' wrapText="1"';
    if (alignment.shrinkToFit) xml += ' shrinkToFit="1"';
    if (alignment.indent !== undefined) xml += ` indent="${alignment.indent}"`;
    if (alignment.textRotation !== undefined) xml += ` textRotation="${alignment.textRotation}"`;
    if (alignment.justifyLastLine) xml += ' justifyLastLine="1"';
    if (alignment.readingOrder !== undefined) xml += ` readingOrder="${alignment.readingOrder}"`;
    return `${xml}/>`;
}

function protectionXml(protection: ProtectionStyle): string {
    let xml = '<protection';
    if (protection.locked === false) xml += ' locked="0"';
    if (protection.hidden) xml += ' hidden="1"';
    return `${xml}/>`;
}

/**
 * Differential style for conditional formatting. Fills use the background
 * color, which is how Excel stores conditional-format highlighting.
 */
function dxfXml(style: CellStyle, numFmtId: number | null): string {
    let xml = '<dxf>';
    if (style.font) xml += fontXml(style.font);
    if (style.numFmt !== undefined && numFmtId !== null) {
        xml += `<numFmt numFmtId="${numFmtId}" formatCode="${escapeXlsxAttribute(style.numFmt)}"/>`;
    }
    if (style.fill) {
        const color = style.fill.bgColor ?? style.fill.fgColor;
        xml += '<fill><patternFill>';
        if (color !== undefined) xml += colorXml(color).replace('<color', '<bgColor');
        xml += '</patternFill></fill>';
    }
    if (style.alignment) xml += alignmentXml(style.alignment);
    if (style.border) xml += borderXml(style.border);
    if (style.protection) xml += protectionXml(style.protection);
    xml += '</dxf>';
    return xml;
}

/** Deduplicating style registry for the XLSX `styles.xml` part. */
export class XlsxStyleRegistry {
    private readonly numFmts = new Map<string, number>();
    private readonly fontIds = new Map<string, number>();
    private readonly fillIds = new Map<string, number>();
    private readonly borderIds = new Map<string, number>();
    private readonly xfIds = new Map<string, number>();
    private readonly fonts: string[] = [BASE_FONT_NORMAL, BASE_FONT_BOLD];
    private readonly fills: string[] = [BASE_FILL_NONE, BASE_FILL_GRAY, BASE_FILL_SOLID];
    private readonly borders: string[] = [BASE_BORDER];
    private readonly xfs: string[] = [...BASE_XFS];
    private readonly styles: (CellStyle | null)[] = [null, null, null, null, null, null];
    private readonly dxfIds = new Map<string, number>();
    private readonly dxfs: string[] = [];
    private nextNumFmtId = 165;

    private numFmtIdFor(format: string): number {
        let id = this.numFmts.get(format);
        if (id === undefined) {
            id = this.nextNumFmtId++;
            this.numFmts.set(format, id);
        }
        return id;
    }

    /** Cell XF index for a style; 0 for null/empty styles. */
    register(style: CellStyle | null | undefined): number {
        const normalized = normalizeStyle(style);
        if (normalized === null) return 0;
        const key = JSON.stringify(normalized);
        const existing = this.xfIds.get(key);
        if (existing !== undefined) return existing;

        let numFmtId = 0;
        if (normalized.numFmt !== undefined) {
            numFmtId = this.numFmtIdFor(normalized.numFmt);
        }

        let fontId = 0;
        if (normalized.font) {
            if (normalized.font.bold && Object.keys(normalized.font).length === 1) {
                fontId = 1;
            } else {
                const componentKey = JSON.stringify(normalized.font);
                let id = this.fontIds.get(componentKey);
                if (id === undefined) {
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
            if (id === undefined) {
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
            if (id === undefined) {
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
        const children = (normalized.alignment ? alignmentXml(normalized.alignment) : '')
            + (normalized.protection ? protectionXml(normalized.protection) : '');
        if (children) {
            xf += `>${children}</xf>`;
        } else {
            xf += '/>';
        }

        const index = this.xfs.length;
        this.xfIds.set(key, index);
        this.xfs.push(xf);
        this.styles.push(normalized);
        return index;
    }

    /** Differential style id for conditional formatting; -1 for an empty style. */
    registerDxf(style: CellStyle | null | undefined): number {
        const normalized = normalizeStyle(style);
        if (normalized === null) return -1;
        const key = JSON.stringify(normalized);
        const existing = this.dxfIds.get(key);
        if (existing !== undefined) return existing;

        const numFmtId = normalized.numFmt !== undefined
            ? this.numFmtIdFor(normalized.numFmt)
            : null;
        const index = this.dxfs.length;
        this.dxfIds.set(key, index);
        this.dxfs.push(dxfXml(normalized, numFmtId));
        return index;
    }

    /** XF index for a plain number format (legacy FormattedCell path). */
    registerFormat(format: string): number {
        return this.register({ numFmt: format });
    }

    toXml(): string {
        let numFmts = '';
        for (const [format, id] of this.numFmts) {
            numFmts += `<numFmt numFmtId="${id}" formatCode="${escapeXlsxAttribute(format)}"/>`;
        }
        const numFmtsSection = this.numFmts.size > 0
            ? `<numFmts count="${this.numFmts.size}">${numFmts}</numFmts>`
            : '';
        const dxfsSection = this.dxfs.length > 0
            ? `<dxfs count="${this.dxfs.length}">${this.dxfs.join('\n')}</dxfs>`
            : '<dxfs count="0"/>';

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${numFmtsSection}
<fonts count="${this.fonts.length}">
${this.fonts.join('\n')}
</fonts>
<fills count="${this.fills.length}">
${this.fills.join('\n')}
</fills>
<borders count="${this.borders.length}">
${this.borders.join('\n')}
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="${this.xfs.length}">
${this.xfs.join('\n')}
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
${dxfsSection}
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
    }
}

export type SharedStringValue = string | RichText;

function sharedStringKey(value: SharedStringValue): string {
    if (typeof value === 'string') return `s:${value}`;
    return `r:${JSON.stringify(value.runs)}`;
}

function richTextXml(rich: RichText): string {
    let xml = '';
    for (const run of rich.runs) {
        xml += '<r>';
        const font = normalizeFont(run.style?.font);
        if (font) {
            xml += fontXml(font).replace('<font>', '<rPr>').replace('</font>', '</rPr>');
        }
        xml += `<t xml:space="preserve">${escapeXlsxText(run.text)}</t></r>`;
    }
    return xml;
}

function sharedStringItem(value: SharedStringValue): string {
    if (typeof value !== 'string') return `<si>${richTextXml(value)}</si>`;
    const clean = escapeXlsxText(value);
    if (clean.length > 0 &&
        (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean))) {
        return `<si><t xml:space="preserve">${clean}</t></si>`;
    }
    return `<si><t>${clean}</t></si>`;
}

export function buildXlsxSharedStrings(strings: SharedStringValue[], countAll: number): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${countAll}" uniqueCount="${strings.length}">`;
    for (const value of strings) xml += sharedStringItem(value);
    xml += '</sst>';
    return xml;
}

export function buildXlsxWorkbook(sheets: XlsxSheetEntry[], autofilterOn: boolean): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
<workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>
<sheets>`;

    for (const sheet of sheets) {
        const state = sheet.hidden ? ' state="hidden"' : '';
        xml += `<sheet name="${escapeXlsxAttribute(sheet.name)}" sheetId="${sheet.sheetId}"${state} r:id="${sheet.rId}"/>`;
    }
    xml += '</sheets>';

    if (autofilterOn) {
        xml += '<definedNames>';
        for (const sheet of sheets) {
            if (!sheet.filterHeaderRange) continue;
            const escapedRange = escapeXlsxText(sheet.filterHeaderRange);
            xml += `<definedName name="_xlnm._FilterDatabase" localSheetId="${sheet.sheetId - 1}" hidden="1">${escapedRange}</definedName>`;
        }
        xml += '</definedNames>';
    }

    xml += '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>';
    return xml;
}

export function buildXlsxContentTypes(
    sheets: XlsxSheetEntry[],
    comments?: Array<{ path: string }>,
): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${comments && comments.length > 0 ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : ''}
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`;
    for (const sheet of sheets) {
        xml += `<Override PartName="/${sheet.pathInArchive}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }
    for (const part of comments ?? []) {
        xml += `<Override PartName="/${part.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
    }
    xml += '</Types>';
    return xml;
}

export function buildXlsxRels(sheets: XlsxSheetEntry[]): { global: string; workbook: string } {
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
    workbook += '</Relationships>';
    return { global, workbook };
}

/**
 * Write everything up to and including `<sheetData>`. The caller writes the
 * header row and data rows, then calls {@link writeXlsxSheetTail}.
 */
export function writeXlsxSheetHead(
    sink: TextByteSink,
    options: {
        dimensionRef: string;
        columnCount: number;
        isFirstSheet: boolean;
        hasHeaders: boolean;
        colWidths: number[];
    },
): void {
    const { dimensionRef, columnCount, isFirstSheet, hasHeaders, colWidths } = options;

    sink.writeString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    sink.writeString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    sink.writeString(`<dimension ref="${dimensionRef}"/>`);

    if (hasHeaders) {
        sink.writeString(isFirstSheet
            ? '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
            : '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" /><selection pane="bottomLeft" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
    } else {
        sink.writeString(isFirstSheet
            ? '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
            : '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>');
    }

    sink.writeString('<cols>');
    for (let i = 0; i < columnCount; i++) {
        const width = defaultColWidth(colWidths[i]);
        sink.writeString(`<col min="${i + 1}" max="${i + 1}" width="${formatColumnWidth(width)}" bestFit="1" customWidth="1" />`);
    }
    sink.writeString('</cols><sheetData>');
}

/**
 * Finish the sheet. Returns the autofilter range for the workbook's defined
 * name, or null when no autofilter was written.
 */
export function writeXlsxSheetTail(
    sink: TextByteSink,
    options: {
        doAutofilter: boolean;
        columnCount: number;
        lastColumnLetter: string;
        rowNum: number;
        sheetName: string;
        merges?: string[];
        hyperlinks?: Array<{ spec: HyperlinkSpec; rId?: string }>;
        dataValidations?: DataValidationSpec[];
        conditionalFormats?: ConditionalFormatSpec[];
        legacyDrawingRId?: string;
        cfPriorityStart?: number;
        /** Register a differential style for conditional formatting. */
        registerDxf?: (style: CellStyle) => number;
    },
): string | null {
    const {
        doAutofilter, columnCount, lastColumnLetter, rowNum, sheetName,
        merges, hyperlinks, dataValidations, conditionalFormats,
        legacyDrawingRId, cfPriorityStart = 1, registerDxf,
    } = options;
    sink.writeString('</sheetData>');

    let filterRange: string | null = null;
    if (doAutofilter && columnCount > 0) {
        const ref = `A1:${lastColumnLetter}${rowNum}`;
        sink.writeString(`<autoFilter ref="${ref}"/>`);
        filterRange = `${formatSheetNameForFormula(sheetName)}!$A$1:$${lastColumnLetter}$${rowNum}`;
    }

    if (merges && merges.length > 0) {
        sink.writeString(`<mergeCells count="${merges.length}">`);
        for (const ref of merges) sink.writeString(`<mergeCell ref="${escapeXmlText(ref)}"/>`);
        sink.writeString('</mergeCells>');
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

    sink.writeString('</worksheet>');
    return filterRange;
}

function columnLetterFromRef(ref: string): { col: number; row: number } {
    const match = /^([A-Za-z]{1,3})(\d+)$/.exec(ref.trim());
    if (!match) throw new Error(`Invalid cell reference '${ref}'`);
    let col = 0;
    for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { col: col - 1, row: parseInt(match[2], 10) - 1 };
}

export function buildXlsxComments(comments: CommentSpec[]): string {
    const authors: string[] = [];
    const authorIds = new Map<string, number>();
    for (const comment of comments) {
        const author = comment.author ?? '';
        if (!authorIds.has(author)) {
            authorIds.set(author, authors.length);
            authors.push(author);
        }
    }
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>';
    for (const author of authors) xml += `<author>${escapeXlsxText(author)}</author>`;
    xml += '</authors><commentList>';
    for (const comment of comments) {
        const text = escapeXlsxText(comment.text);
        const space = text.length > 0 && (text[0] === ' ' || text[text.length - 1] === ' ' || /[\t\n\r]/.test(text))
            ? ' xml:space="preserve"'
            : '';
        xml += `<comment ref="${escapeXmlText(comment.ref)}" authorId="${authorIds.get(comment.author ?? '')}"><text><t${space}>${text}</t></text></comment>`;
    }
    xml += '</commentList></comments>';
    return xml;
}

export function buildXlsxCommentsVml(comments: CommentSpec[]): string {
    const V = 'urn:schemas-microsoft-com:vml';
    const O = 'urn:schemas-microsoft-com:office:office';
    const X = 'urn:schemas-microsoft-com:office:excel';
    let xml = `<xml xmlns:v="${V}" xmlns:o="${O}" xmlns:x="${X}">`;
    xml += `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">`
        + '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
    let nextId = 1025;
    for (const comment of comments) {
        const { col, row } = columnLetterFromRef(comment.ref);
        xml += `<v:shape id="_x0000_s${nextId++}" type="#_x0000_t202" `
            + 'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" '
            + 'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>'
            + '<v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>'
            + '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>'
            + `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill>`
            + `<x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;
    }
    xml += '</xml>';
    return xml;
}

/** Worksheet relationships part for hyperlinks and the comment VML drawing. */
export function buildXlsxSheetRels(
    rels: Array<{ id: string; type: string; target: string; external?: boolean }>,
): string {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (const rel of rels) {
        xml += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXlsxAttribute(rel.target)}"`;
        if (rel.external) xml += ' TargetMode="External"';
        xml += '/>';
    }
    xml += '</Relationships>';
    return xml;
}

export const XLSX_REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
export const XLSX_REL_VML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
export const XLSX_REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export function writeXlsxSharedStringCell(
    sink: TextByteSink,
    sst: { values: SharedStringValue[]; index: Map<string, number> },
    value: SharedStringValue,
    ref: string,
    styleOverride?: number,
): void {
    const key = sharedStringKey(value);
    let index = sst.index.get(key);
    if (index === undefined) {
        index = sst.values.length;
        sst.values.push(value);
        sst.index.set(key, index);
    }
    const styleAttr = styleOverride !== undefined ? ` s="${styleOverride}"` : '';
    sink.writeString(`<c${ref} t="s"${styleAttr}><v>${index}</v></c>`);
}

/** Write one inline string cell (plain or rich text). */
export function writeXlsxInlineStringCell(
    sink: TextByteSink,
    value: SharedStringValue,
    ref: string,
    styleOverride?: number,
): void {
    const styleAttr = styleOverride !== undefined ? ` s="${styleOverride}"` : '';
    if (typeof value !== 'string') {
        sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is>${richTextXml(value)}</is></c>`);
        return;
    }
    const clean = escapeXlsxText(value);
    if (clean.length > 0 &&
        (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean))) {
        sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is><t xml:space="preserve">${clean}</t></is></c>`);
    } else {
        sink.writeString(`<c${ref} t="inlineStr"${styleAttr}><is><t>${clean}</t></is></c>`);
    }
}

export interface HyperlinkSpec {
    /** Cell reference, e.g. `A1`. */
    ref: string;
    /** External URL (needs a worksheet relationship). */
    target?: string;
    /** Internal location, e.g. `Other!A1`. */
    location?: string;
    tooltip?: string;
    display?: string;
}

export interface DataValidationSpec {
    sqref: string;
    type: 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
    operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'lessThan'
        | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';
    formula1?: string;
    formula2?: string;
    /** For `type: 'list'`: literal values, emitted as a quoted comma list. */
    values?: string[];
    allowBlank?: boolean;
    showDropdown?: boolean;
    showInputMessage?: boolean;
    promptTitle?: string;
    prompt?: string;
    showErrorMessage?: boolean;
    errorTitle?: string;
    error?: string;
    errorStyle?: 'stop' | 'warning' | 'information';
}

export type ConditionalFormatRule =
    | {
        type: 'cellIs';
        operator: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'lessThan'
            | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';
        formula: string[];
        style?: CellStyle;
        stopIfTrue?: boolean;
    }
    | {
        type: 'expression';
        formula: string;
        style?: CellStyle;
        stopIfTrue?: boolean;
    }
    | {
        type: 'colorScale';
        cfvo: Array<{ type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: string }>;
        colors: string[];
    }
    | {
        type: 'dataBar';
        cfvo: Array<{ type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: string }>;
        color: string;
        showValue?: boolean;
    }
    | {
        type: 'iconSet';
        iconSet?: '3Arrows' | '3ArrowsGray' | '3Flags' | '3TrafficLights1' | '3TrafficLights2'
            | '3Signs' | '3Symbols' | '3Symbols2' | '4Arrows' | '4ArrowsGray' | '4RedToBlack'
            | '4Rating' | '4TrafficLights' | '5Arrows' | '5ArrowsGray' | '5Rating' | '5Quarters';
        cfvo?: Array<{ type: 'percent' | 'num' | 'percentile'; value?: string }>;
        showValue?: boolean;
        reverse?: boolean;
    }
    | {
        type: 'top10';
        rank?: number;
        percent?: boolean;
        bottom?: boolean;
        style?: CellStyle;
    }
    | {
        type: 'duplicateValues' | 'uniqueValues';
        style?: CellStyle;
    }
    | {
        type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
        text: string;
        style?: CellStyle;
    };

export interface ConditionalFormatSpec {
    sqref: string;
    rules: ConditionalFormatRule[];
}

export interface CommentSpec {
    ref: string;
    text: string;
    author?: string;
    /** Server-side comment GUID, when the comment came from another system. */
    guid?: string;
}

export interface XlsxSheetFeatures {
    merges?: string[];
    hyperlinks?: HyperlinkSpec[];
    dataValidations?: DataValidationSpec[];
    conditionalFormats?: ConditionalFormatSpec[];
    comments?: CommentSpec[];
}

const DV_TYPES: Record<DataValidationSpec['type'], string> = {
    whole: 'whole', decimal: 'decimal', list: 'list', date: 'date',
    time: 'time', textLength: 'textLength', custom: 'custom',
};

const DV_OPERATORS: Record<NonNullable<DataValidationSpec['operator']>, string> = {
    between: 'between', notBetween: 'notBetween', equal: 'equal', notEqual: 'notEqual',
    lessThan: 'lessThan', lessThanOrEqual: 'lessThanOrEqual',
    greaterThan: 'greaterThan', greaterThanOrEqual: 'greaterThanOrEqual',
};

function writeDataValidations(sink: TextByteSink, validations: DataValidationSpec[]): void {
    if (validations.length === 0) return;
    sink.writeString(`<dataValidations count="${validations.length}">`);
    for (const rule of validations) {
        let attrs = ` type="${DV_TYPES[rule.type]}" sqref="${escapeXmlText(rule.sqref)}"`;
        if (rule.operator !== undefined) attrs += ` operator="${DV_OPERATORS[rule.operator]}"`;
        if (rule.allowBlank) attrs += ' allowBlank="1"';
        if (rule.showDropdown === false) attrs += ' showDropDown="1"';
        if (rule.showInputMessage) attrs += ' showInputMessage="1"';
        if (rule.showErrorMessage) attrs += ' showErrorMessage="1"';
        if (rule.errorStyle !== undefined) attrs += ` errorStyle="${rule.errorStyle}"`;
        if (rule.promptTitle !== undefined) attrs += ` promptTitle="${escapeXlsxAttribute(rule.promptTitle)}"`;
        if (rule.prompt !== undefined) attrs += ` prompt="${escapeXlsxAttribute(rule.prompt)}"`;
        if (rule.errorTitle !== undefined) attrs += ` errorTitle="${escapeXlsxAttribute(rule.errorTitle)}"`;
        if (rule.error !== undefined) attrs += ` error="${escapeXlsxAttribute(rule.error)}"`;
        sink.writeString(`<dataValidation${attrs}>`);
        if (rule.values !== undefined) {
            const list = rule.values.map(value => value.replace(/"/g, '""')).join(',');
            sink.writeString(`<formula1>"${escapeXmlText(list)}"</formula1>`);
        } else if (rule.formula1 !== undefined) {
            sink.writeString(`<formula1>${escapeXmlText(rule.formula1.replace(/^=/, ''))}</formula1>`);
        }
        if (rule.formula2 !== undefined) {
            sink.writeString(`<formula2>${escapeXmlText(rule.formula2.replace(/^=/, ''))}</formula2>`);
        }
        sink.writeString('</dataValidation>');
    }
    sink.writeString('</dataValidations>');
}

function firstCellOfSqref(sqref: string): string {
    const range = sqref.trim().split(/\s+/)[0] ?? 'A1';
    return range.split(':')[0] || 'A1';
}

function writeConditionalFormatting(
    sink: TextByteSink,
    groups: ConditionalFormatSpec[],
    startPriority: number,
    registerDxf?: (style: CellStyle) => number,
): void {
    let priority = startPriority;
    for (const group of groups) {
        sink.writeString(`<conditionalFormatting sqref="${escapeXmlText(group.sqref)}">`);
        const firstCell = firstCellOfSqref(group.sqref);
        for (const rule of group.rules) {
            const current = priority++;
            const dxfIndex = 'style' in rule && rule.style
                ? (registerDxf?.(rule.style) ?? -1)
                : -1;
            const dxfAttr = dxfIndex >= 0 ? ` dxfId="${dxfIndex}"` : '';
            switch (rule.type) {
                case 'cellIs': {
                    let attrs = ` type="cellIs" operator="${rule.operator}" priority="${current}"${dxfAttr}`;
                    if (rule.stopIfTrue) attrs += ' stopIfTrue="1"';
                    sink.writeString(`<cfRule${attrs}>`);
                    for (const formula of rule.formula) {
                        sink.writeString(`<formula>${escapeXmlText(formula.replace(/^=/, ''))}</formula>`);
                    }
                    sink.writeString('</cfRule>');
                    break;
                }
                case 'expression': {
                    let attrs = ` type="expression" priority="${current}"${dxfAttr}`;
                    if (rule.stopIfTrue) attrs += ' stopIfTrue="1"';
                    sink.writeString(`<cfRule${attrs}><formula>${escapeXmlText(rule.formula.replace(/^=/, ''))}</formula></cfRule>`);
                    break;
                }
                case 'colorScale': {
                    sink.writeString(`<cfRule type="colorScale" priority="${current}"><colorScale>`);
                    for (const point of rule.cfvo) {
                        sink.writeString(`<cfvo type="${point.type}"${point.value !== undefined ? ` val="${escapeXmlText(point.value)}"` : ''}/>`);
                    }
                    for (const color of rule.colors) {
                        sink.writeString(`<color rgb="${colorKey(color).slice(4)}"/>`);
                    }
                    sink.writeString('</colorScale></cfRule>');
                    break;
                }
                case 'dataBar': {
                    sink.writeString(`<cfRule type="dataBar" priority="${current}"><dataBar${rule.showValue === false ? ' showValue="0"' : ''}>`);
                    for (const point of rule.cfvo) {
                        sink.writeString(`<cfvo type="${point.type}"${point.value !== undefined ? ` val="${escapeXmlText(point.value)}"` : ''}/>`);
                    }
                    sink.writeString(`<color rgb="${colorKey(rule.color).slice(4)}"/></dataBar></cfRule>`);
                    break;
                }
                case 'iconSet': {
                    let attrs = ` type="iconSet" priority="${current}"`;
                    if (rule.showValue === false) attrs += ' showValue="0"';
                    if (rule.reverse) attrs += ' reverse="1"';
                    sink.writeString(`<cfRule${attrs}><iconSet iconSet="${rule.iconSet ?? '3TrafficLights1'}">`);
                    for (const point of rule.cfvo ?? []) {
                        sink.writeString(`<cfvo type="${point.type}"${point.value !== undefined ? ` val="${escapeXmlText(point.value)}"` : ''}/>`);
                    }
                    sink.writeString('</iconSet></cfRule>');
                    break;
                }
                case 'top10': {
                    let attrs = ` type="top10" priority="${current}"${dxfAttr}`;
                    if (rule.rank !== undefined) attrs += ` rank="${rule.rank}"`;
                    if (rule.percent) attrs += ' percent="1"';
                    if (rule.bottom) attrs += ' bottom="1"';
                    sink.writeString(`<cfRule${attrs}/>`);
                    break;
                }
                case 'duplicateValues':
                case 'uniqueValues':
                    sink.writeString(`<cfRule type="${rule.type}" priority="${current}"${dxfAttr}/>`);
                    break;
                case 'containsText':
                case 'notContainsText':
                case 'beginsWith':
                case 'endsWith': {
                    const operator = rule.type === 'containsText' ? 'containsText'
                        : rule.type === 'notContainsText' ? 'notContains'
                        : rule.type === 'beginsWith' ? 'beginsWith' : 'endsWith';
                    const escaped = rule.text.replace(/"/g, '""');
                    sink.writeString(`<cfRule type="${rule.type}" operator="${operator}" text="${escapeXlsxAttribute(rule.text)}" priority="${current}"${dxfAttr}><formula>${escapeXmlText(`NOT(ISERROR(SEARCH("${escaped}",${firstCell})))`)}</formula></cfRule>`);
                    break;
                }
            }
        }
        sink.writeString('</conditionalFormatting>');
    }
}

function writeHyperlinks(
    sink: TextByteSink,
    links: Array<{ spec: HyperlinkSpec; rId?: string }>,
): void {
    if (links.length === 0) return;
    sink.writeString('<hyperlinks>');
    for (const { spec, rId } of links) {
        let attrs = ` ref="${escapeXmlText(spec.ref)}"`;
        if (rId) attrs += ` r:id="${rId}"`;
        if (spec.location !== undefined) attrs += ` location="${escapeXlsxAttribute(spec.location)}"`;
        if (spec.tooltip !== undefined) attrs += ` tooltip="${escapeXlsxAttribute(spec.tooltip)}"`;
        if (spec.display !== undefined) attrs += ` display="${escapeXlsxAttribute(spec.display)}"`;
        sink.writeString(`<hyperlink${attrs}/>`);
    }
    sink.writeString('</hyperlinks>');
}

export interface XlsxCellContext {
    /** Write a string using the table selected by the writer (shared/inline). */
    writeStringCell(sink: TextByteSink, value: SharedStringValue, ref: string, styleOverride?: number): void;
    /** Register a full cell style and return its cellXfs index. */
    registerStyle(style: CellStyle): number;
    /** Register a plain number format and return its cellXfs index. */
    registerFormat(format: string): number;
    toOaDate(date: Date): number;
}

function writeFormulaCell(
    sink: TextByteSink,
    context: XlsxCellContext,
    cell: FormulaCell,
    ref: string,
    styleAttr: string,
    dateStyle: number,
    dateTimeStyle: number,
): void {
    const formula = cell.formula.trim().replace(/^=/, '');
    if (!formula) throw new Error('FormulaCell.formula must not be empty');
    const cached = cell.value;
    let typeAttr = '';
    let cachedXml = '';

    if (cached instanceof ExcelError) {
        typeAttr = ' t="e"';
        cachedXml = `<v>${escapeXmlText(cached.code)}</v>`;
    } else if (cached instanceof Date) {
        const serial = context.toOaDate(cached);
        // A caller-supplied style already says how the date is shown.
        const effectiveStyle = styleAttr
            || ` s="${dateHasTime(cached) ? dateTimeStyle : dateStyle}"`;
        sink.writeString(`<c${ref}${effectiveStyle}><f>${escapeXmlText(formula)}</f><v>${serial}</v></c>`);
        return;
    } else if (typeof cached === 'number' && Number.isFinite(cached)) {
        cachedXml = `<v>${cached}</v>`;
    } else if (typeof cached === 'boolean') {
        typeAttr = ' t="b"';
        cachedXml = `<v>${cached ? 1 : 0}</v>`;
    } else if (cached !== null && cached !== undefined) {
        typeAttr = ' t="str"';
        cachedXml = `<v>${escapeXmlText(String(cached))}</v>`;
    }

    sink.writeString(`<c${ref}${typeAttr}${styleAttr}><f>${escapeXmlText(formula)}</f>${cachedXml}</c>`);
}

/**
 * Write one cell. Returns the column of the last written cell, which the caller
 * passes back as `prevCol`; null cells write nothing and leave it unchanged.
 */
export function writeXlsxCell(
    sink: TextByteSink,
    context: XlsxCellContext,
    raw: CellValue,
    col: number,
    prevCol: number,
    colLetter: string,
    rowNum: number,
    dateStyle: number = 1,
    dateTimeStyle: number = 2,
): number {
    if (raw === null || raw === undefined) return prevCol;

    const style = getCellStyle(raw);
    const xfIdx = style !== null ? context.registerStyle(style) : -1;
    const styleAttr = xfIdx > 0 ? ` s="${xfIdx}"` : '';
    const ref = cellRefAttr(col, prevCol, colLetter, rowNum);
    const stringStyle = xfIdx > 0 ? xfIdx : undefined;

    const formula = getFormula(raw);
    if (formula !== null) {
        writeFormulaCell(sink, context, raw as FormulaCell, ref, styleAttr, dateStyle, dateTimeStyle);
        return col;
    }

    const value = unwrapCell(raw);
    if (value === null || value === undefined) return prevCol;

    if (isRichText(value)) {
        context.writeStringCell(sink, value, ref, stringStyle);
    } else if (typeof value === 'number') {
        if (Number.isFinite(value)) {
            sink.writeString(`<c${ref}${styleAttr}><v>${value}</v></c>`);
        } else {
            context.writeStringCell(sink, value.toString(), ref, stringStyle);
        }
    } else if (value instanceof ExcelError) {
        sink.writeString(`<c${ref} t="e"${styleAttr}><v>${escapeXmlText(value.code)}</v></c>`);
    } else if (typeof value === 'bigint') {
        context.writeStringCell(sink, value.toString(), ref, stringStyle);
    } else if (typeof value === 'boolean') {
        sink.writeString(`<c${ref} t="b"${styleAttr}><v>${value ? 1 : 0}</v></c>`);
    } else if (value instanceof Date) {
        const oaDate = context.toOaDate(value);
        if (Number.isFinite(oaDate)) {
            const finalStyle = xfIdx > 0
                ? xfIdx
                : (dateHasTime(value) ? dateTimeStyle : dateStyle);
            sink.writeString(`<c${ref} s="${finalStyle}"><v>${oaDate}</v></c>`);
        } else {
            context.writeStringCell(sink, value.toString(), ref, stringStyle);
        }
    } else {
        context.writeStringCell(sink, value.toString(), ref, stringStyle);
    }

    return col;
}
