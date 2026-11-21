/**
 * XLSB conditional formatting: BrtDXF differential styles and the
 * BrtBeginConditionalFormatting / BrtBeginCFRule blocks.
 *
 * The record layouts mirror the ones Excel writes for the same rules saved
 * from an XLSX workbook, so the shared XLSX/XLSB rule API produces visually
 * identical conditional formats in both formats.
 */
import {
    AlignmentStyle,
    BorderStyle,
    CellStyle,
    FillStyle,
    FontStyle,
    StyleColor,
} from '../../Formats';
import { ConditionalFormatRule, ConditionalFormatSpec } from './xlsxParts';
import { buildBrtRecord, concat } from './xlsbRecords';
import { compileXlsbFormula, compileXlsbFormulaStats } from './xlsbFormula';
import { colorKey, parseRangeRef } from './sheetLayout';

const BRT_BEGIN_DXFS = 0x1f9;
const BRT_END_DXFS = 0x1fa;
const BRT_DXF = 0x1fb;
const BRT_BEGIN_CF = 0x1cd;
const BRT_END_CF = 0x1ce;
const BRT_BEGIN_CF_RULE = 0x1cf;
const BRT_END_CF_RULE = 0x1d0;
const BRT_BEGIN_ICON_SET = 0x1d1;
const BRT_END_ICON_SET = 0x1d2;
const BRT_BEGIN_DATA_BAR = 0x1d3;
const BRT_END_DATA_BAR = 0x1d4;
const BRT_BEGIN_COLOR_SCALE = 0x1d5;
const BRT_END_COLOR_SCALE = 0x1d6;
const BRT_CFVO = 0x1d7;
const BRT_COLOR = 0x234;

// CFType / CFTemp ([MS-XLSB] 2.5.17, 2.5.18).
const CF_CELL_IS = 1;
const CF_EXPRESSION = 2;
const CF_TOP10 = 5;
const CF_COLOR_SCALE = 3;
const CF_DATA_BAR = 4;
const CF_ICON_SET = 6;
const TEMPLATE_FORMULA = 0x01;
const TEMPLATE_FILTER = 0x05;
const TEMPLATE_UNIQUE = 0x07;
const TEMPLATE_DUPLICATE = 0x1b;

const CELL_IS_CODES: Record<string, number> = {
    between: 1,
    notBetween: 2,
    equal: 3,
    notEqual: 4,
    greaterThan: 5,
    lessThan: 6,
    greaterThanOrEqual: 7,
    lessThanOrEqual: 8,
};

const TEXT_OPERATORS: Record<string, number> = {
    containsText: 0,
    notContainsText: 1,
    beginsWith: 2,
    endsWith: 3,
};

const CFVO_CODES: Record<string, number> = {
    num: 1,
    min: 2,
    max: 3,
    percent: 4,
    percentile: 5,
    formula: 7,
};

const ICON_SET_CODES: Record<string, number> = {
    '3Arrows': 0,
    '3ArrowsGray': 1,
    '3Flags': 2,
    '3TrafficLights1': 3,
    '3TrafficLights2': 4,
    '3Signs': 5,
    '3Symbols': 6,
    '3Symbols2': 7,
    '4Arrows': 8,
    '4ArrowsGray': 9,
    '4RedToBlack': 10,
    '4Rating': 11,
    '4TrafficLights': 12,
    '5Arrows': 13,
    '5ArrowsGray': 14,
    '5Rating': 15,
    '5Quarters': 16,
};

// XFPropType ([MS-XLSB] 2.5.165).
const PROP_FILL_PATTERN = 0x00;
const PROP_FOREGROUND = 0x01;
const PROP_BACKGROUND = 0x02;
const PROP_FONT_COLOR = 0x05;
const PROP_BORDER_TOP = 0x06;
const PROP_BORDER_BOTTOM = 0x07;
const PROP_BORDER_LEFT = 0x08;
const PROP_BORDER_RIGHT = 0x09;
const PROP_HORIZONTAL = 0x0f;
const PROP_WRAP = 0x14;
const PROP_FONT_NAME = 0x18;
const PROP_BOLD = 0x19;
const PROP_UNDERLINE = 0x1a;
const PROP_ITALIC = 0x1c;
const PROP_STRIKE = 0x1d;
const PROP_SIZE = 0x24;
const PROP_NUMBER_FORMAT = 0x26;
const PROP_NUMBER_FORMAT_ID = 0x29;
const PROP_LOCKED = 0x2b;
const PROP_HIDDEN = 0x2c;

const DXF_FLAGS = 0x8000;
const STOP_IF_TRUE = 0x02;
const PERCENT_FLAG = 0x08;
const BOTTOM_FLAG = 0x10;

const PATTERN_CODES: Record<string, number> = {
    none: 0, solid: 1, mediumGray: 2, darkGray: 3, lightGray: 4,
    darkHorizontal: 5, darkVertical: 6, darkDown: 7, darkUp: 8,
    darkGrid: 9, darkTrellis: 10, lightHorizontal: 11, lightVertical: 12,
    lightDown: 13, lightUp: 14, lightGrid: 15, lightTrellis: 16,
    gray125: 17, gray0625: 18,
};

const BORDER_CODES: Record<string, number> = {
    none: 0x00, thin: 0x01, medium: 0x02, dashed: 0x03, dotted: 0x04,
    thick: 0x05, double: 0x06, hair: 0x07, mediumDashed: 0x08,
    dashDot: 0x09, mediumDashDot: 0x0a, dashDotDot: 0x0b,
    mediumDashDotDot: 0x0c, slantDashDot: 0x0d,
};

const UNDERLINE_CODES: Record<string, number> = {
    single: 0x01,
    double: 0x02,
    singleAccounting: 0x21,
    doubleAccounting: 0x22,
};

const HORIZONTAL_CODES: Record<string, number> = {
    general: 0, left: 1, center: 2, right: 3, fill: 4, justify: 5,
    centerContinuous: 6, distributed: 7,
};

function u16(value: number): number[] {
    const v = value & 0xffff;
    return [v & 0xff, (v >>> 8) & 0xff];
}

function u32(value: number): number[] {
    const v = value >>> 0;
    return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
}

function i32(value: number): number[] {
    return u32(value >>> 0);
}

function doubleBytes(value: number): number[] {
    const buffer = new ArrayBuffer(8);
    new DataView(buffer).setFloat64(0, value, true);
    return [...new Uint8Array(buffer)];
}

function wideString(text: string): number[] {
    const out = [...u32(text.length)];
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        out.push(code & 0xff, (code >>> 8) & 0xff);
    }
    return out;
}

function nullableString(text: string | undefined): number[] {
    return text === undefined ? [0xff, 0xff, 0xff, 0xff] : wideString(text);
}

/** An 8-byte XFPropColor / BrtColor payload. */
function cfColor(color: StyleColor | undefined, opaqueRgb = false): number[] {
    if (color === undefined) return [0, 0, 0, 0, 0, 0, 0, 0];
    const key = colorKey(color);
    if (key === 'auto') return [0x01, 0x40, 0, 0, 0, 0, 0, 0];
    if (key.startsWith('rgb:')) {
        const value = key.slice(4);
        const alpha = opaqueRgb ? 0xff : parseInt(value.slice(0, 2), 16);
        return [
            0x05, 0xff, 0, 0,
            parseInt(value.slice(2, 4), 16),
            parseInt(value.slice(4, 6), 16),
            parseInt(value.slice(6, 8), 16),
            alpha,
        ];
    }
    if (key.startsWith('indexed:')) {
        return [0x03, parseInt(key.slice(8), 10) & 0xff, 0, 0, 0, 0, 0, 0xff];
    }
    const [, themeText, tintText] = key.split(':');
    const tint = Math.max(-32768, Math.min(32767, Math.round(Number(tintText) * 32767)));
    return [
        0x07, parseInt(themeText, 10) & 0xff,
        tint & 0xff, (tint >> 8) & 0xff, 0, 0, 0, 0xff,
    ];
}

function prop(kind: number, data: number[]): { kind: number; bytes: number[] } {
    return {
        kind,
        bytes: [...u16(kind), ...u16(4 + data.length), ...data],
    };
}

function unsupported(name: string): never {
    throw new Error(
        `${name} is not supported in XLSB conditional formats; use XlsxWriter for it`,
    );
}

function fontProps(font: FontStyle): Array<{ kind: number; bytes: number[] }> {
    if (font.vertAlign !== undefined) unsupported('Font.vertAlign');
    if (font.family !== undefined) unsupported('Font.family');
    if (font.scheme !== undefined) unsupported('Font.scheme');
    if (font.outline !== undefined) unsupported('Font.outline');
    if (font.shadow !== undefined) unsupported('Font.shadow');
    if (font.condense !== undefined) unsupported('Font.condense');
    if (font.extend !== undefined) unsupported('Font.extend');
    if (font.charset !== undefined) unsupported('Font.charset');
    const props: Array<{ kind: number; bytes: number[] }> = [];
    if (font.color !== undefined) props.push(prop(PROP_FONT_COLOR, cfColor(font.color)));
    if (font.name !== undefined) props.push(prop(PROP_FONT_NAME, wideString(font.name)));
    if (font.bold !== undefined) props.push(prop(PROP_BOLD, u16(font.bold ? 700 : 400)));
    if (font.underline !== undefined) {
        const key = font.underline === true ? 'single' : font.underline;
        props.push(prop(PROP_UNDERLINE, u16(key === false ? 0 : (UNDERLINE_CODES[key] ?? 0))));
    }
    if (font.italic !== undefined) props.push(prop(PROP_ITALIC, [font.italic ? 1 : 0]));
    if (font.strike !== undefined) props.push(prop(PROP_STRIKE, [font.strike ? 1 : 0]));
    if (font.size !== undefined) props.push(prop(PROP_SIZE, u32(Math.round(font.size * 20))));
    return props;
}

function fillProps(fill: FillStyle): Array<{ kind: number; bytes: number[] }> {
    if (fill.gradient) unsupported('GradientFill');
    const props: Array<{ kind: number; bytes: number[] }> = [];
    if (fill.patternType !== undefined) {
        props.push(prop(PROP_FILL_PATTERN, [PATTERN_CODES[fill.patternType] ?? 0]));
    }
    if (fill.fgColor !== undefined) props.push(prop(PROP_FOREGROUND, cfColor(fill.fgColor)));
    if (fill.bgColor !== undefined) {
        props.push(prop(PROP_BACKGROUND, cfColor(fill.bgColor)));
    } else if (fill.patternType === 'solid' && fill.fgColor !== undefined) {
        // A solid conditional fill paints the cell with its background color.
        props.push(prop(PROP_BACKGROUND, cfColor(fill.fgColor)));
    }
    return props;
}

function borderProps(border: BorderStyle): Array<{ kind: number; bytes: number[] }> {
    for (const side of ['diagonal', 'start', 'end', 'vertical', 'horizontal'] as const) {
        if (border[side] !== undefined) unsupported(`Border.${side}`);
    }
    const props: Array<{ kind: number; bytes: number[] }> = [];
    const sides: Array<[keyof BorderStyle, number]> = [
        ['top', PROP_BORDER_TOP],
        ['bottom', PROP_BORDER_BOTTOM],
        ['left', PROP_BORDER_LEFT],
        ['right', PROP_BORDER_RIGHT],
    ];
    for (const [name, kind] of sides) {
        const side = border[name] as BorderStyle['top'];
        const style = side?.style ?? 'none';
        const color = side?.color === undefined ? [0x01, 0x40, 0, 0, 0, 0, 0, 0] : cfColor(side.color);
        props.push(prop(kind, [...color, ...u16(BORDER_CODES[style] ?? 0)]));
    }
    return props;
}

function alignmentProps(alignment: AlignmentStyle): Array<{ kind: number; bytes: number[] }> {
    for (const name of ['vertical', 'shrinkToFit', 'indent', 'textRotation', 'relativeIndent',
        'justifyLastLine', 'readingOrder'] as const) {
        if (alignment[name] !== undefined) unsupported(`Alignment.${name}`);
    }
    const props: Array<{ kind: number; bytes: number[] }> = [];
    if (alignment.horizontal !== undefined) {
        props.push(prop(PROP_HORIZONTAL, [HORIZONTAL_CODES[alignment.horizontal] ?? 0]));
    }
    if (alignment.wrapText !== undefined) {
        props.push(prop(PROP_WRAP, [alignment.wrapText ? 1 : 0]));
    }
    return props;
}

/** Build a BrtDXF payload from a differential cell style. */
export function buildXlsbDxfPayload(
    style: CellStyle,
    registerNumFmt: (format: string) => number,
): Uint8Array {
    const props: Array<{ kind: number; bytes: number[] }> = [];
    if (style.font) props.push(...fontProps(style.font));
    if (style.fill) props.push(...fillProps(style.fill));
    if (style.border) props.push(...borderProps(style.border));
    if (style.alignment) props.push(...alignmentProps(style.alignment));
    if (style.protection?.locked !== undefined) {
        props.push(prop(PROP_LOCKED, [style.protection.locked ? 1 : 0]));
    }
    if (style.protection?.hidden !== undefined) {
        props.push(prop(PROP_HIDDEN, [style.protection.hidden ? 1 : 0]));
    }
    if (style.numFmt !== undefined) {
        const id = registerNumFmt(style.numFmt);
        props.push({
            kind: PROP_NUMBER_FORMAT - 0.5,
            bytes: [...u16(PROP_NUMBER_FORMAT_ID), ...u16(6), ...u16(id & 0xffff)],
        });
        props.push(prop(PROP_NUMBER_FORMAT, wideString(style.numFmt)));
    }
    props.sort((left, right) => left.kind - right.kind);
    return Uint8Array.from([
        ...u16(DXF_FLAGS), ...u16(0), ...u16(props.length),
        ...props.flatMap(item => item.bytes),
    ]);
}

/** Deduplicating differential-style registry for `xl/styles.bin`. */
export class XlsbDxfRegistry {
    private readonly ids = new Map<string, number>();
    private readonly payloads: Uint8Array[] = [];

    constructor(private readonly registerNumFmt: (format: string) => number) {}

    register(style: CellStyle | null | undefined): number {
        if (!style) return -1;
        const key = JSON.stringify(style);
        const existing = this.ids.get(key);
        if (existing !== undefined) return existing;
        const id = this.payloads.length;
        this.ids.set(key, id);
        this.payloads.push(buildXlsbDxfPayload(style, this.registerNumFmt));
        return id;
    }

    get size(): number {
        return this.payloads.length;
    }

    /** BrtBeginDXFs / BrtDXF* / BrtEndDXFs, for insertion in styles.bin. */
    toRecords(): Uint8Array {
        if (this.payloads.length === 0) return new Uint8Array(0);
        const parts: Uint8Array[] = [
            buildBrtRecord(BRT_BEGIN_DXFS, Uint8Array.from(u32(this.payloads.length))),
        ];
        for (const payload of this.payloads) {
            parts.push(buildBrtRecord(BRT_DXF, payload));
        }
        parts.push(buildBrtRecord(BRT_END_DXFS, new Uint8Array(0)));
        return concat(parts);
    }
}

function cfvoRecord(
    type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula',
    value: string | undefined,
    gte: boolean,
    origin: { col: number; row: number },
): Uint8Array {
    const code = CFVO_CODES[type];
    if (type === 'formula') {
        if (value === undefined) throw new Error('A formula CFVO needs a value');
        const compiled = compileXlsbFormulaStats(value, { origin, relativeOffsets: true });
        const size = compiled.tokens.length + 2 * compiled.references;
        return buildBrtRecord(BRT_CFVO, Uint8Array.from([
            ...u32(code), ...doubleBytes(0), ...u32(gte ? 1 : 0), ...u32(gte ? 1 : 0),
            ...u32(size), ...u32(compiled.tokens.length), ...compiled.tokens, ...u32(0),
        ]));
    }
    const number = type === 'min' || type === 'max' ? 0 : Number(value ?? 0);
    if (!Number.isFinite(number)) throw new RangeError('CFVO values must be finite');
    return buildBrtRecord(BRT_CFVO, Uint8Array.from([
        ...u32(code), ...doubleBytes(number), ...u32(gte ? 1 : 0), ...u32(gte ? 1 : 0),
        ...u32(0),
    ]));
}

function colorRecord(color: string): Uint8Array {
    return buildBrtRecord(BRT_COLOR, Uint8Array.from(cfColor(color, true)));
}

function rulePayload(
    kind: number,
    template: number,
    dxfId: number,
    priority: number,
    param: number,
    flags: number,
    formulas: string[],
    text: string | undefined,
    origin: { col: number; row: number },
): Uint8Array {
    const compiled = formulas.map(formula =>
        compileXlsbFormulaStats(formula, { origin, relativeOffsets: true }));
    const sizes = compiled.map(item => item.tokens.length + 2 * item.references);
    while (sizes.length < 3) sizes.push(0);
    const payload: number[] = [
        ...u32(kind), ...u32(template), ...i32(dxfId), ...u32(priority), ...u32(param),
        ...u32(0), ...u32(0),
        ...u16(flags), ...u32(sizes[0]), ...u32(sizes[1]), ...u32(sizes[2]),
        ...nullableString(text),
    ];
    for (const item of compiled) {
        // CellParsedFormula: token count, tokens, extra-data length (zero).
        payload.push(...u32(item.tokens.length), ...item.tokens, ...u32(0));
    }
    return Uint8Array.from(payload);
}

function scaleRule(
    rule: ConditionalFormatRule & { type: 'colorScale' | 'dataBar' | 'iconSet' },
    priority: number,
    origin: { col: number; row: number },
): Uint8Array {
    const kind = rule.type === 'colorScale' ? CF_COLOR_SCALE
        : rule.type === 'dataBar' ? CF_DATA_BAR : CF_ICON_SET;
    const template = kind === CF_COLOR_SCALE ? 0x02 : kind === CF_DATA_BAR ? 0x03 : 0x04;
    const head = buildBrtRecord(BRT_BEGIN_CF_RULE, Uint8Array.from([
        ...u32(kind), ...u32(template), ...i32(-1), ...u32(priority),
        ...u32(0), ...u32(0), ...u32(0), ...u16(0), ...u32(0), ...u32(0), ...u32(0),
        ...i32(-1),
    ]));

    if (rule.type === 'colorScale') {
        if (rule.colors.length < 2) throw new Error('A color scale needs at least two colors');
        const parts: Uint8Array[] = [buildBrtRecord(BRT_BEGIN_COLOR_SCALE, new Uint8Array(0))];
        const cfvo = rule.cfvo;
        if (cfvo.length === 2) {
            parts.push(cfvoRecord(cfvo[0].type, cfvo[0].value, false, origin));
            parts.push(cfvoRecord(cfvo[1].type, cfvo[1].value, false, origin));
        } else {
            const first = cfvo[0] ?? { type: 'min' as const };
            const middle = cfvo[1] ?? { type: 'percentile' as const, value: '50' };
            const last = cfvo[2] ?? { type: 'max' as const };
            parts.push(cfvoRecord(first.type, first.value, false, origin));
            parts.push(cfvoRecord(middle.type, middle.value, false, origin));
            parts.push(cfvoRecord(last.type, last.value, false, origin));
        }
        for (const color of rule.colors) parts.push(colorRecord(color));
        parts.push(buildBrtRecord(BRT_END_COLOR_SCALE, new Uint8Array(0)));
        return concat([head, ...parts, buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0))]);
    }

    if (rule.type === 'dataBar') {
        const cfvo = rule.cfvo;
        const first = cfvo[0] ?? { type: 'min' as const };
        const last = cfvo[1] ?? { type: 'max' as const };
        return concat([
            head,
            buildBrtRecord(BRT_BEGIN_DATA_BAR, Uint8Array.from([
                10, 90, rule.showValue === false ? 0 : 1,
            ])),
            cfvoRecord(first.type, first.value, false, origin),
            cfvoRecord(last.type, last.value, false, origin),
            colorRecord(rule.color),
            buildBrtRecord(BRT_END_DATA_BAR, new Uint8Array(0)),
            buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0)),
        ]);
    }

    const iconSet = rule.iconSet ?? '3TrafficLights1';
    const iconCode = ICON_SET_CODES[iconSet];
    if (iconCode === undefined) throw new Error(`Unknown XLSB icon set '${iconSet}'`);
    const count = iconSet.startsWith('5') ? 5 : iconSet.startsWith('4') ? 4 : 3;
    const defaults = Array.from({ length: count - 1 }, (_, index) => ({
        type: 'percent' as const,
        value: String(Math.round(((index + 1) * 100) / count)),
    }));
    const supplied = rule.cfvo ?? [];
    const thresholds = supplied.length >= count - 1
        ? supplied.slice(0, count - 1)
        : [...supplied, ...defaults.slice(supplied.length)];
    // XLSB stores one BrtCFVO per icon; the first one is ignored by Excel but
    // must be present. The public API (like XLSX) supplies only the N-1 real
    // thresholds, so prepend the 0% baseline.
    const values = [
        { type: 'percent' as const, value: '0' },
        ...thresholds,
    ];
    const flags = (rule.showValue === false ? 0x02 : 0) | (rule.reverse ? 0x04 : 0);
    const parts: Uint8Array[] = [
        head,
        buildBrtRecord(BRT_BEGIN_ICON_SET, Uint8Array.from([...u32(iconCode), ...u16(flags)])),
    ];
    for (const item of values) parts.push(cfvoRecord(item.type, item.value, true, origin));
    parts.push(buildBrtRecord(BRT_END_ICON_SET, new Uint8Array(0)));
    parts.push(buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0)));
    return concat(parts);
}

function textFormula(rule: {
    type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith';
    text: string;
}, corner: string): string {
    const quoted = `"${rule.text.replace(/"/g, '""')}"`;
    switch (rule.type) {
        case 'containsText':
            return `NOT(ISERROR(SEARCH(${quoted},${corner})))`;
        case 'notContainsText':
            return `ISERROR(SEARCH(${quoted},${corner}))`;
        case 'beginsWith':
            return `LEFT(${corner},${rule.text.length})=${quoted}`;
        case 'endsWith':
            return `RIGHT(${corner},${rule.text.length})=${quoted}`;
    }
}

function ruleBin(
    rule: ConditionalFormatRule,
    sqref: string,
    priority: number,
    dxfRegistry: XlsbDxfRegistry,
): Uint8Array {
    const start = parseRangeRef(sqref.trim().split(/\s+/)[0]);
    const origin = { col: start.c1, row: start.r1 };
    if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
        return scaleRule(rule, priority, origin);
    }

    const corner = `${columnLetter(start.c1)}${start.r1 + 1}`;
    const dxfId = dxfRegistry.register(rule.style);
    const stop = 'stopIfTrue' in rule && rule.stopIfTrue ? STOP_IF_TRUE : 0;

    let payload: Uint8Array;
    if (rule.type === 'cellIs') {
        const operator = CELL_IS_CODES[rule.operator];
        if (operator === undefined) {
            throw new Error(`CellIs operator '${rule.operator}' is not supported in XLSB`);
        }
        payload = rulePayload(CF_CELL_IS, 0, dxfId, priority, operator, stop,
            rule.formula, undefined, origin);
    } else if (rule.type === 'expression') {
        payload = rulePayload(CF_EXPRESSION, TEMPLATE_FORMULA, dxfId, priority, 0, stop,
            [rule.formula], undefined, origin);
    } else if (rule.type === 'top10') {
        payload = rulePayload(
            CF_TOP10, TEMPLATE_FILTER, dxfId, priority, rule.rank ?? 10,
            stop | (rule.percent ? PERCENT_FLAG : 0) | (rule.bottom ? BOTTOM_FLAG : 0),
            [], undefined, origin,
        );
    } else if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
        payload = rulePayload(
            CF_EXPRESSION,
            rule.type === 'uniqueValues' ? TEMPLATE_UNIQUE : TEMPLATE_DUPLICATE,
            dxfId, priority, 0, stop, [], undefined, origin,
        );
    } else if (rule.type === 'containsText' || rule.type === 'notContainsText'
        || rule.type === 'beginsWith' || rule.type === 'endsWith') {
        const template = TEXT_OPERATORS[rule.type];
        payload = rulePayload(
            CF_EXPRESSION, 0x08, dxfId, priority, template, stop,
            [textFormula(rule, corner)], rule.text, origin,
        );
    } else {
        throw new Error(`Unsupported XLSB conditional format rule '${(rule as { type: string }).type}'`);
    }
    return concat([
        buildBrtRecord(BRT_BEGIN_CF_RULE, payload),
        buildBrtRecord(BRT_END_CF_RULE, new Uint8Array(0)),
    ]);
}

function columnLetter(index: number): string {
    let value = index + 1;
    let result = '';
    while (value > 0) {
        const remainder = (value - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        value = Math.floor((value - 1) / 26);
    }
    return result;
}

/** BrtBeginConditionalFormatting / rules / BrtEndConditionalFormatting blocks. */
export function buildXlsbConditionalFormatRecords(
    formats: ConditionalFormatSpec[],
    dxfRegistry: XlsbDxfRegistry,
    startPriority: number = 1,
): Uint8Array {
    const parts: Uint8Array[] = [];
    let priority = startPriority;
    for (const format of formats) {
        if (!format.rules || format.rules.length === 0) continue;
        const ranges = format.sqref.trim().split(/\s+/).map(parseRangeRef);
        if (ranges.length === 0) throw new Error('Conditional formatting sqref must not be empty');
        const head: number[] = [
            ...u32(format.rules.length), ...u32(0), ...u32(ranges.length),
        ];
        for (const range of ranges) {
            head.push(...u32(range.r1), ...u32(range.r2), ...u32(range.c1), ...u32(range.c2));
        }
        parts.push(buildBrtRecord(BRT_BEGIN_CF, Uint8Array.from(head)));
        for (const rule of format.rules) {
            parts.push(ruleBin(rule, format.sqref, priority++, dxfRegistry));
        }
        parts.push(buildBrtRecord(BRT_END_CF, new Uint8Array(0)));
    }
    return concat(parts);
}

/** Number of priorities used by the given formats, for multi-sheet ordering. */
export function countConditionalFormatRules(formats: ConditionalFormatSpec[] | undefined): number {
    return (formats ?? []).reduce((count, format) => count + format.rules.length, 0);
}

// Keep the raw compiler import referenced so bundlers do not drop it; the
// stats variant is the only one used here.
void compileXlsbFormula;
