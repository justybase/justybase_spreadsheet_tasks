/**
 * An Excel error value (`#DIV/0!`, `#N/A`, ...). Strings are always text, so
 * writing an error requires this marker; readers return the display text
 * (`"#DIV/0!"`) for compatibility with plain value consumers.
 */
export class ExcelError {
    constructor(public readonly code: string) {}

    toString(): string {
        return this.code;
    }
}

export type PrimitiveCellValue = string | number | boolean | Date | bigint | ExcelError | null | undefined;

/** RGB/ARGB hex (`'FF0000'`, `'#FF0000'`, `'FFFF0000'`) or an indexed/theme color. */
export type StyleColor = string | { indexed: number } | { theme: number; tint?: number };

export type FontUnderline = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export interface FontStyle {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: FontUnderline | boolean;
    strike?: boolean;
    color?: StyleColor;
    family?: number;
    scheme?: 'major' | 'minor';
    vertAlign?: 'baseline' | 'superscript' | 'subscript';
    outline?: boolean;
    shadow?: boolean;
    condense?: boolean;
    extend?: boolean;
    charset?: number;
}

export type PatternType =
    | 'none' | 'solid' | 'mediumGray' | 'darkGray' | 'lightGray'
    | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid'
    | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown'
    | 'lightUp' | 'lightGrid' | 'lightTrellis' | 'gray125' | 'gray0625';

export interface FillStyle {
    patternType?: PatternType;
    fgColor?: StyleColor;
    bgColor?: StyleColor;
}

export type BorderLineStyle =
    | 'thin' | 'medium' | 'dashed' | 'dotted' | 'thick' | 'double' | 'hair'
    | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot'
    | 'mediumDashDotDot' | 'slantDashDot';

export interface BorderSide {
    style?: BorderLineStyle;
    color?: StyleColor;
}

export interface BorderStyle {
    top?: BorderSide;
    bottom?: BorderSide;
    left?: BorderSide;
    right?: BorderSide;
    diagonal?: BorderSide;
    diagonalUp?: boolean;
    diagonalDown?: boolean;
    outline?: boolean;
}

export interface AlignmentStyle {
    horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
    vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
    wrapText?: boolean;
    shrinkToFit?: boolean;
    indent?: number;
    textRotation?: number;
    justifyLastLine?: boolean;
    readingOrder?: 0 | 1 | 2;
}

export interface ProtectionStyle {
    locked?: boolean;
    hidden?: boolean;
}

/** Full cell style. Only the properties that are set are applied. */
export interface CellStyle {
    numFmt?: string;
    font?: FontStyle;
    fill?: FillStyle;
    border?: BorderStyle;
    alignment?: AlignmentStyle;
    protection?: ProtectionStyle;
}

export interface StyledCell {
    value: PrimitiveCellValue | RichText;
    style: CellStyle;
}

/** A formula and an optional cached value. The formula text may start with `=`. */
export interface FormulaCell {
    formula: string;
    value?: PrimitiveCellValue;
    style?: CellStyle;
}

export interface RichTextRun {
    text: string;
    /** Run formatting; only font properties are used. */
    style?: CellStyle;
}

/** Rich text split into formatting runs. */
export class RichText {
    constructor(public readonly runs: RichTextRun[]) {}

    toString(): string {
        return this.runs.map(run => run.text).join('');
    }
}

export interface FormattedCell {
    value: PrimitiveCellValue;
    format: string;
}

export type CellValue =
    | PrimitiveCellValue
    | RichText
    | FormattedCell
    | StyledCell
    | FormulaCell;

/** Built-in Excel number-format strings for {@link FormattedCell}. */
export const F = {
    THOUSANDS_SEP: '#,##0',
    CURRENCY_PLN: '#,##0.00 "z\u0142"',
    CURRENCY_EUR: '#,##0.00 \u20AC',
    PERCENTAGE: '0%',
    SCIENTIFIC: '0.00E+00',
    TWO_DECIMALS: '#,##0.00',
    TEXT: '@',
    LEADING_ZEROS: '000000000',

    DATE_SHORT: 'dd.mm.yyyy',
    DATE_LONG: 'd mmmm yyyy',
    DATE_DAY_MONTH_YEAR: 'dd-mm-yyyy',
    DATE_ISO: 'yyyy-mm-dd',
    DATE_MONTH_YEAR: 'mmmm yyyy',
    DATE_WEEKDAY: 'dddd, d mmmm yyyy',
    DATE_DAY_MONTH: 'd mmmm',
    DATE_YEAR_ONLY: 'yyyy',

    DATETIME_SHORT: 'dd.mm.yyyy hh:mm',
    DATETIME_LONG: 'd mmmm yyyy hh:mm:ss',
    TIME_HH_MM: 'hh:mm',
    TIME_HH_MM_SS: 'hh:mm:ss',
    TIME_12H: 'h:mm AM/PM',
    DATETIME_24H: 'dd.mm.yyyy hh:mm:ss',
    DATETIME_ISO: 'yyyy-mm-dd"T"hh:mm:ss',
    TIME_MS: 'hh:mm:ss.000',
};

export function isFormattedCell(val: unknown): val is FormattedCell {
    return val !== null && val !== undefined && typeof val === 'object' && 'format' in val && typeof (val as FormattedCell).format === 'string';
}

export function isRichText(val: unknown): val is RichText {
    return val instanceof RichText;
}

export function isStyledCell(val: unknown): val is StyledCell {
    return val !== null && val !== undefined && typeof val === 'object'
        && 'style' in val && 'value' in val && !('formula' in val);
}

export function isFormulaCell(val: unknown): val is FormulaCell {
    return val !== null && val !== undefined && typeof val === 'object' && 'formula' in val;
}

/**
 * Style carried by a cell. `FormattedCell` maps to `{ numFmt }`; every other
 * value has no style.
 */
export function getCellStyle(val: CellValue): CellStyle | null {
    if (isFormattedCell(val)) return { numFmt: val.format };
    if (isStyledCell(val)) return val.style;
    if (isFormulaCell(val)) return val.style ?? null;
    return null;
}

/** Style with the legacy `FormattedCell.format` merged into `numFmt`. */
export function mergeCellStyle(base: CellStyle | null | undefined, extra: CellStyle | null | undefined): CellStyle | null {
    if (!base && !extra) return null;
    return { ...(base ?? {}), ...(extra ?? {}) };
}

export function getFormula(val: CellValue): string | null {
    return isFormulaCell(val) ? val.formula : null;
}

export function unwrapCell(val: CellValue): PrimitiveCellValue | RichText {
    if (isFormattedCell(val) || isStyledCell(val)) {
        return val.value;
    }
    if (isFormulaCell(val)) {
        return val.value ?? null;
    }
    return val as PrimitiveCellValue | RichText;
}

export function getFormat(val: CellValue): string | null {
    const style = getCellStyle(val);
    return style?.numFmt ?? null;
}
