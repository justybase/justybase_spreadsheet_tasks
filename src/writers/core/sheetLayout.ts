/**
 * Sheet layout and range-style helpers shared by the XLSX and XLSB writers.
 *
 * Rows and columns are zero-based internally, like the rest of the writer
 * code; A1 references are only used at the API boundary (freeze panes, tables,
 * range styles).
 */
import { CellStyle, StyleColor, mergeCellStyles } from '../../Formats';
import { columnIndexToLetter } from '../../xmlUtils';

export const MAX_EXCEL_ROW = 1_048_575;
export const MAX_EXCEL_COL = 16_383;

export interface GroupSpec {
    /** First row (0-based) or column letter/index (0-based). */
    start: number | string;
    /** Last row (0-based) or column letter/index (0-based), inclusive. */
    end: number | string;
    collapsed?: boolean;
}

export interface NormalizedGroup {
    start: number;
    end: number;
    collapsed: boolean;
}

export interface Outline {
    level: Uint8Array;
    hidden: Uint8Array;
    collapsed: Uint8Array;
    maxLevel: number;
    /** Indices covered by any group or marked as a summary cell. */
    indices: number[];
}

export interface TableColumnSpec {
    name: string;
    totalsRowFunction?:
        | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'sum'
        | 'stdDev' | 'var' | 'custom';
    totalsRowLabel?: string;
    /** Formula used when `totalsRowFunction` is `custom`. */
    totalsRowFormula?: string;
}

export interface TableSpec {
    /** Unique table name; Excel displays it as the table name. */
    name: string;
    /** A1 range including the header row (and the totals row when present). */
    ref: string;
    columns?: Array<string | TableColumnSpec>;
    headerRow?: boolean;
    totalsRow?: boolean;
    styleName?: string;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
}

export interface SheetProtectionSpec {
    enabled?: boolean;
    /** Legacy Excel password; written as its 16-bit hash. */
    password?: string;
    selectLockedCells?: boolean;
    selectUnlockedCells?: boolean;
    formatCells?: boolean;
    formatColumns?: boolean;
    formatRows?: boolean;
    insertColumns?: boolean;
    insertRows?: boolean;
    insertHyperlinks?: boolean;
    deleteColumns?: boolean;
    deleteRows?: boolean;
    sort?: boolean;
    autoFilter?: boolean;
    pivotTables?: boolean;
    objects?: boolean;
    scenarios?: boolean;
}

export interface CellRange {
    c1: number;
    r1: number;
    c2: number;
    r2: number;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/;
const OPEN_COLUMNS_RE = /^\$?([A-Za-z]{1,3})(?:\$?([1-9]\d*))?:\$?([A-Za-z]{1,3})$/;
const OPEN_ROWS_RE = /^\$?([1-9]\d*):\$?([1-9]\d*)$/;

export function columnIndex(letters: string): number {
    let value = 0;
    for (const ch of letters.toUpperCase()) {
        const code = ch.charCodeAt(0);
        if (code < 65 || code > 90) throw new Error(`Invalid column '${letters}'`);
        value = value * 26 + (code - 64);
    }
    return value - 1;
}

export function parseCellRef(ref: string): { row: number; col: number } {
    const match = CELL_RE.exec(ref.trim());
    if (!match) throw new Error(`Invalid cell reference '${ref}'`);
    const col = columnIndex(match[1]);
    const row = parseInt(match[2], 10) - 1;
    if (col > MAX_EXCEL_COL || row > MAX_EXCEL_ROW) {
        throw new RangeError(`Cell reference '${ref}' is outside the Excel grid`);
    }
    return { row, col };
}

/** Parse `A1:B2`, a single cell, or an empty range (A1). */
export function parseRangeRef(ref: string): CellRange {
    const parts = ref.trim().split(':');
    if (parts.length > 2) throw new Error(`Invalid range '${ref}'`);
    const first = parseCellRef(parts[0]);
    const last = parts.length === 2 ? parseCellRef(parts[1]) : first;
    return {
        c1: Math.min(first.col, last.col),
        r1: Math.min(first.row, last.row),
        c2: Math.max(first.col, last.col),
        r2: Math.max(first.row, last.row),
    };
}

/** Parse a `styles` key: A1 range, whole columns (`B:D`) or whole rows (`1:3`). */
export function parseStyleRange(ref: string): CellRange {
    const text = ref.trim();
    if (!text) throw new Error('Style range must not be empty');

    const columns = OPEN_COLUMNS_RE.exec(text);
    if (columns) {
        const c1 = columnIndex(columns[1]);
        const c2 = columnIndex(columns[3]);
        const r1 = columns[2] === undefined ? 0 : parseInt(columns[2], 10) - 1;
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

export interface NormalizedStyleRange extends CellRange {
    style: CellStyle;
}

/** A row-oriented sweep over the caller's style ranges. */
export class StyleRangeSet {
    private readonly ranges: NormalizedStyleRange[];
    private _row = -1;
    private _active: NormalizedStyleRange[] = [];

    constructor(styles?: Record<string, CellStyle> | NormalizedStyleRange[]) {
        if (!styles) {
            this.ranges = [];
            return;
        }
        if (Array.isArray(styles)) {
            this.ranges = styles;
            return;
        }
        this.ranges = Object.entries(styles).map(([ref, style]) => {
            if (!style || typeof style !== 'object') {
                throw new TypeError(`Style for '${ref}' must be a CellStyle object`);
            }
            return { ...parseStyleRange(ref), style };
        });
    }

    get length(): number {
        return this.ranges.length;
    }

    /** Ranges covering `rowIndex`, in the caller's order. */
    active(rowIndex: number): NormalizedStyleRange[] {
        if (rowIndex === this._row) return this._active;
        this._row = rowIndex;
        this._active = this.ranges.filter(range => range.r1 <= rowIndex && rowIndex <= range.r2);
        return this._active;
    }

    /** Merge every range style covering the cell, in order. */
    styleFor(rowIndex: number, columnIndex: number): CellStyle | null {
        let style: CellStyle | null = null;
        for (const range of this.active(rowIndex)) {
            if (range.c1 <= columnIndex && columnIndex <= range.c2) {
                style = mergeCellStyles(style, range.style);
            }
        }
        return style;
    }
}

function normalizeGroupValue(value: number | string, kind: 'row' | 'column'): number {
    if (typeof value === 'number') {
        if (!Number.isInteger(value) || value < 0) {
            throw new RangeError(`Group ${kind} must be a non-negative integer, got ${value}`);
        }
        return value;
    }
    if (kind === 'row') {
        const parsed = parseInt(value, 10);
        if (!Number.isInteger(parsed) || parsed < 1) {
            throw new RangeError(`Row group bounds must be 1-based row numbers, got '${value}'`);
        }
        return parsed - 1;
    }
    return columnIndex(value);
}

export function normalizeGroups(
    groups: GroupSpec[] | undefined,
    kind: 'row' | 'column',
): NormalizedGroup[] {
    if (!groups) return [];
    return groups.map(group => {
        const start = normalizeGroupValue(group.start, kind);
        const end = normalizeGroupValue(group.end, kind);
        if (end < start) throw new RangeError(`Group end precedes start in ${kind} group`);
        const limit = kind === 'row' ? MAX_EXCEL_ROW : MAX_EXCEL_COL;
        if (end > limit) throw new RangeError(`${kind} group is outside the Excel grid`);
        return { start, end, collapsed: group.collapsed === true };
    });
}

/** Build outline levels, hidden flags and summary markers for one axis. */
export function buildOutline(groups: NormalizedGroup[], maxIndex: number): Outline {
    const size = Math.max(0, maxIndex + 1);
    const level = new Uint8Array(size);
    const hidden = new Uint8Array(size);
    const collapsed = new Uint8Array(size);
    const indices = new Set<number>();

    for (let i = 0; i < groups.length; i++) {
        for (let j = i + 1; j < groups.length; j++) {
            const a = groups[i];
            const b = groups[j];
            const same = a.start === b.start && a.end === b.end;
            if (same) throw new RangeError('Duplicate row/column group');
            const disjoint = a.end < b.start || b.end < a.start;
            const nested = (a.start <= b.start && b.end <= a.end)
                || (b.start <= a.start && a.end <= b.end);
            if (!disjoint && !nested) {
                throw new RangeError('Row/column groups may nest or be disjoint, not overlap partially');
            }
        }
    }

    for (const group of groups) {
        const depth = groups.filter(other =>
            other !== group && other.start <= group.start && group.end <= other.end,
        ).length + 1;
        if (depth > 7) throw new RangeError('Excel supports at most seven outline levels');
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

export interface FreezePane {
    xSplit: number;
    ySplit: number;
    topLeftCell: string;
    activePane: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
    frozen: boolean;
}

/**
 * Resolve the `freezePanes` option. `undefined` keeps the historical default
 * (freeze the first header row when headers are present); `false` disables the
 * freeze; a cell such as `B2` freezes everything above and to its left.
 */
export function resolveFreezePanes(
    value: string | false | null | undefined,
    hasHeaders: boolean,
): FreezePane | null {
    let ref: string | null;
    if (value === undefined) ref = hasHeaders ? 'A2' : null;
    else if (value === false || value === null || value === '') ref = null;
    else ref = value;
    if (ref === null) return null;

    const { row, col } = parseCellRef(ref);
    if (row === 0 && col === 0) return null;
    const activePane = row > 0 && col > 0 ? 'bottomRight'
        : col > 0 ? 'topRight'
            : 'bottomLeft';
    return {
        xSplit: col,
        ySplit: row,
        topLeftCell: `${columnIndexToLetter(col)}${row + 1}`,
        activePane,
        frozen: true,
    };
}

export function validateZoom(value: number | undefined): number | undefined {
    if (value === undefined) return undefined;
    if (!Number.isInteger(value) || value < 10 || value > 400) {
        throw new RangeError(`zoom must be an integer 10..400, got ${value}`);
    }
    return value;
}

export function validateRowHeight(value: number): number {
    if (!Number.isFinite(value) || value <= 0 || value > 409.5) {
        throw new RangeError(`Row height must be in (0, 409.5] points, got ${value}`);
    }
    return value;
}

export function validateColumnWidthValue(value: number): number {
    if (!Number.isFinite(value) || value < 0.5 || value > 255) {
        throw new RangeError(`Column width must be 0.5..255 characters, got ${value}`);
    }
    return value;
}

/** Excel's legacy XOR password hash used by sheet protection. */
export function passwordHash(password: string): number {
    // Excel's legacy 16-bit sheet-protection hash: walk the password
    // backwards, rotate and XOR each byte, then mix in the length and the
    // 0xCE4B constant. (The OOXML prose has the constant/order wrong; this is
    // the algorithm Excel actually uses.)
    let hash = 0;
    for (let i = password.length - 1; i >= 0; i--) {
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
        hash ^= password.charCodeAt(i) & 0xff;
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
    hash ^= 0xce4b;
    hash ^= password.length;
    return hash & 0xffff;
}

export interface NormalizedTableColumn extends TableColumnSpec {
    name: string;
}

export interface NormalizedTableSpec {
    name: string;
    ref: CellRange;
    refText: string;
    columns: NormalizedTableColumn[];
    headerRow: boolean;
    totalsRow: boolean;
    styleName: string;
    showFirstColumn: boolean;
    showLastColumn: boolean;
    showRowStripes: boolean;
    showColumnStripes: boolean;
}

const TABLE_NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;
const A1_REFERENCE_RE = /^([A-Za-z]{1,3})([1-9]\d*)$/;
const R1C1_REFERENCE_RE = /^R[1-9]\d*C[1-9]\d*$/i;

function tableNameIsCellReference(name: string): boolean {
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

export function normalizeTableSpec(spec: TableSpec, id: number): NormalizedTableSpec {
    if (!spec || typeof spec !== 'object') throw new TypeError('tables must be TableSpec objects');
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
    const columns = (spec.columns ?? []).map(column => {
        if (typeof column === 'string') return { name: column };
        if (!column || typeof column.name !== 'string' || !column.name) {
            throw new Error(`Table '${name}' has a column without a name`);
        }
        return { ...column };
    });
    if (columns.length === 0) {
        throw new Error(`Table '${name}' needs columns (pass TableSpec.columns)`);
    }
    if (columns.length !== width) {
        throw new Error(
            `Table '${name}' has ${columns.length} columns but ref '${spec.ref}' is ${width} columns wide`,
        );
    }
    const names = new Set<string>();
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
        styleName: spec.styleName ?? 'TableStyleMedium9',
        showFirstColumn: spec.showFirstColumn === true,
        showLastColumn: spec.showLastColumn === true,
        showRowStripes: spec.showRowStripes !== false,
        showColumnStripes: spec.showColumnStripes === true,
    };
}

export function tableAutoFilterRef(table: NormalizedTableSpec): string | null {
    if (!table.headerRow) return null;
    const endRow = table.totalsRow ? table.ref.r2 - 1 : table.ref.r2;
    if (endRow < table.ref.r1) return null;
    return `${columnIndexToLetter(table.ref.c1)}${table.ref.r1 + 1}:`
        + `${columnIndexToLetter(table.ref.c2)}${endRow + 1}`;
}

/**
 * SUBTOTAL codes for the built-in table totals functions. Excel rejects a
 * table whose totals function and totals-row cell formula disagree, so the
 * writers generate the matching cell formula themselves.
 */
const TOTALS_SUBTOTAL_CODE: Record<string, number> = {
    average: 101,
    count: 103,
    countNums: 102,
    max: 104,
    min: 105,
    sum: 109,
    stdDev: 107,
    var: 108,
};

export interface TableTotalsOverride {
    /** Absolute zero-based worksheet row of the totals cell. */
    row: number;
    /** Zero-based column index. */
    col: number;
    /** Formula without the leading `=`. */
    formula?: string;
    /** Text for a label cell. */
    text?: string;
}

/**
 * The totals-row cells a table definition implies. Built-in functions get the
 * SUBTOTAL formula Excel expects in the cell; custom totals use their formula;
 * labels become text. Undefined functions leave the caller's cell untouched.
 */
export function tableTotalsOverrides(tables: NormalizedTableSpec[]): TableTotalsOverride[] {
    const overrides: TableTotalsOverride[] = [];
    for (const table of tables) {
        if (!table.totalsRow) continue;
        for (let index = 0; index < table.columns.length; index++) {
            const column = table.columns[index];
            const col = table.ref.c1 + index;
            const fn = column.totalsRowFunction;
            if (fn === 'custom') {
                const formula = column.totalsRowFormula?.trim().replace(/^=/, '');
                if (!formula) {
                    throw new Error(
                        `Table '${table.name}' column '${column.name}' needs totalsRowFormula for a custom total`,
                    );
                }
                overrides.push({ row: table.ref.r2, col, formula });
            } else if (fn !== undefined) {
                const code = TOTALS_SUBTOTAL_CODE[fn];
                const firstDataRow = table.headerRow ? table.ref.r1 + 1 : table.ref.r1;
                const lastDataRow = table.ref.r2 - 1;
                const letter = columnIndexToLetter(col);
                if (lastDataRow < firstDataRow) {
                    throw new Error(
                        `Table '${table.name}' needs at least one data row before its totals row`,
                    );
                }
                overrides.push({
                    row: table.ref.r2,
                    col,
                    formula: `SUBTOTAL(${code},${letter}${firstDataRow + 1}:${letter}${lastDataRow + 1})`,
                });
            } else if (column.totalsRowLabel !== undefined) {
                overrides.push({ row: table.ref.r2, col, text: column.totalsRowLabel });
            }
        }
    }
    return overrides;
}

export function tableColumnName(index: number): string {
    return `Column${index + 1}`;
}

/** Color helpers for the writers (auto/indexed/theme/RGB). */
export function colorKey(color: StyleColor): string {
    if (typeof color === 'string') {
        let value = color.trim().replace(/^#/, '').toUpperCase();
        if (value.length === 6) value = `FF${value}`;
        if (!/^[0-9A-F]{8}$/.test(value)) {
            throw new Error(`Invalid color '${color}'. Use RRGGBB, AARRGGBB or '#' + hex.`);
        }
        return `rgb:${value}`;
    }
    if ('indexed' in color) return `indexed:${color.indexed}`;
    if ('auto' in color) return 'auto';
    return `theme:${color.theme}:${color.tint ?? 0}`;
}
