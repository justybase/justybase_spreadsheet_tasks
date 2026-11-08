/**
 * A deliberately small XLSB formula compiler.
 *
 * Supports literals, A1 references/ranges, unary/binary operators and a
 * curated set of worksheet functions. Anything outside this subset fails with
 * a clear error instead of writing a formula Excel cannot parse.
 */
import { ExcelError } from '../../Formats';
import { excelErrorToBiffCode } from '../../biff12Utils';
import { buildBrtRecord } from './xlsbRecords';

const BRT_FMLA_STRING = 0x0008;
const BRT_FMLA_NUM = 0x0009;
const BRT_FMLA_BOOL = 0x000a;
const BRT_FMLA_ERROR = 0x000b;

type OperandClass = 'V' | 'R' | 'A';

interface FunctionInfo {
    id: number;
    min: number;
    max: number;
    fixed: number | null;
    classes: string;
}

const FUNCTIONS: Record<string, FunctionInfo> = {
    COUNT: { id: 0, min: 0, max: 255, fixed: null, classes: 'R' },
    IF: { id: 1, min: 2, max: 3, fixed: null, classes: 'VRR' },
    ISNA: { id: 2, min: 1, max: 1, fixed: 1, classes: 'V' },
    ISERROR: { id: 3, min: 1, max: 1, fixed: 1, classes: 'V' },
    SUM: { id: 4, min: 0, max: 255, fixed: null, classes: 'R' },
    AVERAGE: { id: 5, min: 1, max: 255, fixed: null, classes: 'R' },
    MIN: { id: 6, min: 1, max: 255, fixed: null, classes: 'R' },
    MAX: { id: 7, min: 1, max: 255, fixed: null, classes: 'R' },
    ROW: { id: 8, min: 0, max: 1, fixed: null, classes: 'R' },
    COLUMN: { id: 9, min: 0, max: 1, fixed: null, classes: 'R' },
    NA: { id: 10, min: 0, max: 0, fixed: 0, classes: '' },
    STDEV: { id: 12, min: 1, max: 255, fixed: null, classes: 'R' },
    SIN: { id: 15, min: 1, max: 1, fixed: 1, classes: 'V' },
    COS: { id: 16, min: 1, max: 1, fixed: 1, classes: 'V' },
    TAN: { id: 17, min: 1, max: 1, fixed: 1, classes: 'V' },
    ATAN: { id: 18, min: 1, max: 1, fixed: 1, classes: 'V' },
    PI: { id: 19, min: 0, max: 0, fixed: 0, classes: '' },
    SQRT: { id: 20, min: 1, max: 1, fixed: 1, classes: 'V' },
    EXP: { id: 21, min: 1, max: 1, fixed: 1, classes: 'V' },
    LN: { id: 22, min: 1, max: 1, fixed: 1, classes: 'V' },
    LOG10: { id: 23, min: 1, max: 1, fixed: 1, classes: 'V' },
    ABS: { id: 24, min: 1, max: 1, fixed: 1, classes: 'V' },
    INT: { id: 25, min: 1, max: 1, fixed: 1, classes: 'V' },
    SIGN: { id: 26, min: 1, max: 1, fixed: 1, classes: 'V' },
    ROUND: { id: 27, min: 2, max: 2, fixed: 2, classes: 'VV' },
    INDEX: { id: 29, min: 2, max: 4, fixed: null, classes: 'RVVV' },
    MID: { id: 31, min: 3, max: 3, fixed: 3, classes: 'VVV' },
    LEN: { id: 32, min: 1, max: 1, fixed: 1, classes: 'V' },
    VALUE: { id: 33, min: 1, max: 1, fixed: 1, classes: 'V' },
    AND: { id: 36, min: 1, max: 255, fixed: null, classes: 'R' },
    OR: { id: 37, min: 1, max: 255, fixed: null, classes: 'R' },
    NOT: { id: 38, min: 1, max: 1, fixed: 1, classes: 'V' },
    MOD: { id: 39, min: 2, max: 2, fixed: 2, classes: 'VV' },
    VAR: { id: 46, min: 1, max: 255, fixed: null, classes: 'R' },
    TEXT: { id: 48, min: 2, max: 2, fixed: 2, classes: 'VV' },
    MATCH: { id: 64, min: 2, max: 3, fixed: null, classes: 'VRR' },
    DATE: { id: 65, min: 3, max: 3, fixed: 3, classes: 'VVV' },
    DAY: { id: 67, min: 1, max: 1, fixed: 1, classes: 'V' },
    MONTH: { id: 68, min: 1, max: 1, fixed: 1, classes: 'V' },
    YEAR: { id: 69, min: 1, max: 1, fixed: 1, classes: 'V' },
    HOUR: { id: 71, min: 1, max: 1, fixed: 1, classes: 'V' },
    MINUTE: { id: 72, min: 1, max: 1, fixed: 1, classes: 'V' },
    SECOND: { id: 73, min: 1, max: 1, fixed: 1, classes: 'V' },
    NOW: { id: 74, min: 0, max: 0, fixed: 0, classes: '' },
    ROWS: { id: 76, min: 1, max: 1, fixed: 1, classes: 'R' },
    COLUMNS: { id: 77, min: 1, max: 1, fixed: 1, classes: 'R' },
    SEARCH: { id: 82, min: 2, max: 3, fixed: null, classes: 'VVV' },
    HLOOKUP: { id: 101, min: 3, max: 4, fixed: null, classes: 'VRRV' },
    VLOOKUP: { id: 102, min: 3, max: 4, fixed: null, classes: 'VRRV' },
    LOG: { id: 109, min: 1, max: 2, fixed: null, classes: 'VV' },
    LOWER: { id: 112, min: 1, max: 1, fixed: 1, classes: 'V' },
    UPPER: { id: 113, min: 1, max: 1, fixed: 1, classes: 'V' },
    LEFT: { id: 115, min: 1, max: 2, fixed: null, classes: 'VV' },
    RIGHT: { id: 116, min: 1, max: 2, fixed: null, classes: 'VV' },
    EXACT: { id: 117, min: 2, max: 2, fixed: 2, classes: 'VV' },
    TRIM: { id: 118, min: 1, max: 1, fixed: 1, classes: 'V' },
    REPLACE: { id: 119, min: 4, max: 4, fixed: 4, classes: 'VVVV' },
    SUBSTITUTE: { id: 120, min: 3, max: 4, fixed: null, classes: 'VVVV' },
    FIND: { id: 124, min: 2, max: 3, fixed: null, classes: 'VVV' },
    ISTEXT: { id: 127, min: 1, max: 1, fixed: 1, classes: 'V' },
    ISNUMBER: { id: 128, min: 1, max: 1, fixed: 1, classes: 'V' },
    ISBLANK: { id: 129, min: 1, max: 1, fixed: 1, classes: 'V' },
    T: { id: 130, min: 1, max: 1, fixed: 1, classes: 'R' },
    N: { id: 131, min: 1, max: 1, fixed: 1, classes: 'R' },
    TRUNC: { id: 197, min: 1, max: 2, fixed: null, classes: 'VV' },
    ROUNDUP: { id: 212, min: 2, max: 2, fixed: 2, classes: 'VV' },
    ROUNDDOWN: { id: 213, min: 2, max: 2, fixed: 2, classes: 'VV' },
    RANK: { id: 216, min: 2, max: 3, fixed: null, classes: 'VRV' },
    TODAY: { id: 221, min: 0, max: 0, fixed: 0, classes: '' },
    SUMPRODUCT: { id: 228, min: 1, max: 255, fixed: null, classes: 'A' },
    FLOOR: { id: 285, min: 2, max: 2, fixed: 2, classes: 'VV' },
    CEILING: { id: 288, min: 2, max: 2, fixed: 2, classes: 'VV' },
    CONCATENATE: { id: 336, min: 0, max: 255, fixed: null, classes: 'V' },
    POWER: { id: 337, min: 2, max: 2, fixed: 2, classes: 'VV' },
    SUMIF: { id: 345, min: 2, max: 3, fixed: null, classes: 'RVR' },
    COUNTIF: { id: 346, min: 2, max: 2, fixed: 2, classes: 'RV' },
    COUNTBLANK: { id: 347, min: 1, max: 1, fixed: 1, classes: 'R' },
    AVERAGEA: { id: 361, min: 1, max: 255, fixed: null, classes: 'R' },
    MAXA: { id: 362, min: 1, max: 255, fixed: null, classes: 'R' },
    MINA: { id: 363, min: 1, max: 255, fixed: null, classes: 'R' },
    IFERROR: { id: 480, min: 2, max: 2, fixed: 2, classes: 'VV' },
    COUNTIFS: { id: 481, min: 2, max: 255, fixed: null, classes: 'RV' },
    SUMIFS: { id: 482, min: 3, max: 255, fixed: null, classes: 'RRV' },
    AVERAGEIF: { id: 483, min: 2, max: 3, fixed: null, classes: 'RVR' },
};

const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT', 'CELL', 'INFO']);

const OPERATOR_IDS: Record<string, number> = {
    '+': 0x03, '-': 0x04, '*': 0x05, '/': 0x06, '^': 0x07, '&': 0x08,
    '<': 0x09, '<=': 0x0a, '=': 0x0b, '>=': 0x0c, '>': 0x0d, '<>': 0x0e,
    ' ': 0x0f, ',': 0x10, ':': 0x11,
};

const PRECEDENCE: Record<string, number> = {
    ',': 1, '=': 2, '<': 2, '>': 2, '<=': 2, '>=': 2, '<>': 2,
    '&': 3, '+': 4, '-': 4, '*': 5, '/': 5, '^': 6, ' ': 9, ':': 10,
};

const ERROR_CODES: Record<string, number> = {
    '#NULL!': 0x00,
    '#DIV/0!': 0x07,
    '#VALUE!': 0x0f,
    '#REF!': 0x17,
    '#NAME?': 0x1d,
    '#NUM!': 0x24,
    '#N/A': 0x2a,
    '#GETTING_DATA': 0x2b,
};

const ADDRESS = String.raw`(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)`;

const TOKEN_RE = new RegExp(
    String.raw`\s+|"(?:[^"]|"")*"|'(?:[^']|'')*'`
    + String.raw`|#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)`
    + `|${ADDRESS}(?![\\w.])`
    + String.raw`|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[^\W\d][\w.]*|<>|<=|>=|[+\-*/^&%=<>(),:]`,
    'giu',
);

type Node =
    | { kind: 'number'; value: number }
    | { kind: 'string'; value: string }
    | { kind: 'bool'; value: boolean }
    | { kind: 'error'; value: number }
    | { kind: 'missing' }
    | { kind: 'reference'; value: string }
    | { kind: 'paren'; value: Node }
    | { kind: 'percent'; value: Node }
    | { kind: 'unary'; op: '+' | '-'; value: Node }
    | { kind: 'binary'; op: string; left: Node; right: Node }
    | { kind: 'function'; name: string; args: Node[] };

function tokenize(text: string): string[] {
    const tokens: string[] = [];
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

class Parser {
    private index = 0;

    constructor(private readonly tokens: string[]) {}

    private peek(): string | undefined {
        return this.tokens[this.index];
    }

    private pop(): string {
        const token = this.tokens[this.index++];
        if (token === undefined) throw new Error('Unexpected end of XLSB formula');
        return token;
    }

    parse(): Node {
        const node = this.expression(0);
        if (this.peek() !== undefined) {
            throw new Error(`Unsupported XLSB formula token '${this.peek()}'`);
        }
        return node;
    }

    private expression(minPrecedence: number): Node {
        let left = this.primary();
        while (this.peek() === '%') {
            this.pop();
            left = { kind: 'percent', value: left };
        }
        for (;;) {
            const operator = this.peek();
            if (operator === undefined) break;
            const precedence = PRECEDENCE[operator];
            if (precedence === undefined || precedence < minPrecedence) break;
            this.pop();
            const right = this.expression(precedence + 1);
            left = { kind: 'binary', op: operator, left, right };
        }
        return left;
    }

    private primary(): Node {
        const token = this.pop();
        if (token === '+' || token === '-') {
            return { kind: 'unary', op: token, value: this.expression(7) };
        }
        if (token === '(') {
            const value = this.expression(0);
            if (this.pop() !== ')') throw new Error('Missing ) in XLSB formula');
            return { kind: 'paren', value };
        }
        if (token.startsWith('"')) {
            return { kind: 'string', value: token.slice(1, -1).replace(/""/g, '"') };
        }
        if (token.startsWith('#')) {
            const code = ERROR_CODES[token.toUpperCase()];
            if (code === undefined) throw new Error(`Unsupported XLSB formula error literal '${token}'`);
            return { kind: 'error', value: code };
        }
        if (/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(token)) {
            return { kind: 'number', value: Number(token) };
        }
        if (/^[A-Za-z_\\]/.test(token)) {
            const upper = token.toUpperCase();
            if (this.peek() === '(') {
                this.pop();
                const args: Node[] = [];
                if (this.peek() !== ')') {
                    for (;;) {
                        args.push(this.peek() === ',' || this.peek() === ')'
                            ? { kind: 'missing' }
                            : this.expression(2));
                        if (this.peek() !== ',') break;
                        this.pop();
                    }
                }
                if (this.pop() !== ')') throw new Error('Missing ) in XLSB function call');
                return { kind: 'function', name: upper, args };
            }
            if (upper === 'TRUE' || upper === 'FALSE') {
                return { kind: 'bool', value: upper === 'TRUE' };
            }
            if (ADDRESS_RE.test(token)) {
                return { kind: 'reference', value: token };
            }
            throw new Error(`Unsupported XLSB formula name '${token}' (defined names are not supported)`);
        }
        if (ADDRESS_RE.test(token)) {
            return { kind: 'reference', value: token };
        }
        throw new Error(`Unsupported XLSB formula token '${token}'`);
    }
}

const ADDRESS_RE = new RegExp(`^(?:${ADDRESS})$`, 'i');

function columnIndex(letters: string): number {
    let value = 0;
    for (const ch of letters.toUpperCase()) {
        value = value * 26 + (ch.charCodeAt(0) - 64);
    }
    return value - 1;
}

interface ReferencePart {
    row: number;
    col: number;
    rowRelative: boolean;
    colRelative: boolean;
}

function parseReferencePart(text: string, kind: 'cell' | 'col' | 'row'): ReferencePart {
    if (kind === 'cell') {
        const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(text);
        if (!match) throw new Error(`Invalid XLSB cell reference '${text}'`);
        const col = columnIndex(match[2]);
        const row = parseInt(match[4], 10) - 1;
        if (col > 16383 || row > 1048575) {
            throw new RangeError(`Cell reference '${text}' is outside the Excel grid`);
        }
        return { row, col, colRelative: match[1] === '', rowRelative: match[3] === '' };
    }
    if (kind === 'col') {
        const match = /^(\$?)([A-Za-z]{1,3})$/.exec(text);
        if (!match) throw new Error(`Invalid XLSB column reference '${text}'`);
        return { row: 0, col: columnIndex(match[2]), rowRelative: false, colRelative: match[1] === '' };
    }
    const match = /^(\$?)([1-9]\d*)$/.exec(text);
    if (!match) throw new Error(`Invalid XLSB row reference '${text}'`);
    return { row: parseInt(match[2], 10) - 1, col: 0, rowRelative: match[1] === '', colRelative: false };
}

function refKind(part: string): 'cell' | 'col' | 'row' {
    if (/^[A-Za-z]/.test(part.replace(/^\$/, ''))) {
        return /^\d|^\$\d/.test(part.replace(/^\$?[A-Za-z]{1,3}/, '')) ? 'cell' : 'col';
    }
    return 'row';
}

function columnField(part: ReferencePart): number {
    return (part.col & 0x3fff)
        | (part.colRelative ? 0x4000 : 0)
        | (part.rowRelative ? 0x8000 : 0);
}

function emitReference(text: string, operand: OperandClass): number[] {
    const parts = text.split(':');
    const base = operand === 'R' ? 0x20 : operand === 'A' ? 0x60 : 0x40;
    if (parts.length === 1) {
        const part = parseReferencePart(parts[0], 'cell');
        return [base | 0x04, ...u32(part.row), ...u16(columnField(part))];
    }
    if (parts.length !== 2) throw new Error(`Invalid XLSB range '${text}'`);
    const kind = refKind(parts[0]);
    if (kind !== refKind(parts[1])) {
        throw new Error(`Invalid XLSB range '${text}'`);
    }
    let first = parseReferencePart(parts[0], kind);
    let last = parseReferencePart(parts[1], kind);
    if (kind === 'col') {
        first = { ...first, row: 0, rowRelative: false };
        last = { ...last, row: 1048575, rowRelative: false };
    } else if (kind === 'row') {
        first = { ...first, col: 0, colRelative: false };
        last = { ...last, col: 16383, colRelative: false };
    }
    return [
        base | 0x05,
        ...u32(first.row), ...u32(last.row),
        ...u16(columnField(first)), ...u16(columnField(last)),
    ];
}

function u16(value: number): number[] {
    const v = value & 0xffff;
    return [v & 0xff, (v >>> 8) & 0xff];
}

function u32(value: number): number[] {
    const v = value >>> 0;
    return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
}

function doubleBytes(value: number): number[] {
    const buffer = new Uint8Array(8);
    new DataView(buffer.buffer).setFloat64(0, value, true);
    return [...buffer];
}

function stringBytes(text: string): number[] {
    if (text.length > 255) {
        throw new RangeError('XLSB formula string literals are limited to 255 UTF-16 units');
    }
    const out = [...u16(text.length)];
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        out.push(code & 0xff, (code >>> 8) & 0xff);
    }
    return out;
}

function functionArgumentClass(info: FunctionInfo, index: number): OperandClass {
    if (info.id === 481) return index % 2 === 0 ? 'R' : 'V';
    if (info.id === 482 || info.id === 484) {
        return index === 0 || index % 2 === 0 ? 'R' : 'V';
    }
    const chars = info.classes;
    if (chars.length === 0) return 'V';
    return chars[Math.min(index, chars.length - 1)] === 'R' ? 'R' : 'V';
}

function emitNode(node: Node, operand: OperandClass = 'V'): number[] {
    switch (node.kind) {
        case 'number': {
            if (Number.isInteger(node.value) && node.value >= 0 && node.value <= 65535) {
                return [0x1e, ...u16(node.value)];
            }
            return [0x1f, ...doubleBytes(node.value)];
        }
        case 'string':
            return [0x17, ...stringBytes(node.value)];
        case 'bool':
            return [0x1d, node.value ? 1 : 0];
        case 'error':
            return [0x1c, node.value];
        case 'missing':
            return [0x16];
        case 'reference':
            return emitReference(node.value, operand);
        case 'paren':
            return [...emitNode(node.value, operand), 0x15];
        case 'percent':
            return [...emitNode(node.value), 0x14];
        case 'unary':
            return [...emitNode(node.value), node.op === '+' ? 0x12 : 0x13];
        case 'binary': {
            const argType: OperandClass = node.op === ',' || node.op === ' ' || node.op === ':'
                ? 'R'
                : 'V';
            return [
                ...emitNode(node.left, argType),
                ...emitNode(node.right, argType),
                OPERATOR_IDS[node.op],
            ];
        }
        case 'function': {
            const info = FUNCTIONS[node.name];
            if (!info) {
                throw new Error(`Unsupported XLSB function '${node.name}'`);
            }
            if (node.args.length < info.min || node.args.length > info.max) {
                throw new RangeError(
                    `${node.name} expects ${info.min}..${info.max} arguments, got ${node.args.length}`,
                );
            }
            const encoded = node.args.map((arg, index) => emitNode(arg, functionArgumentClass(info, index)));
            const resultClass = operand === 'R' ? 0x20 : operand === 'A' ? 0x60 : 0x40;
            const ending = info.fixed !== null
                ? [resultClass | 0x01, ...u16(info.id)]
                : [resultClass | 0x02, node.args.length, ...u16(info.id)];

            if (info.id === 1) {
                const yes = encoded[1] ?? [];
                const no = encoded[2] ?? [];
                const firstGoto = [0x19, 0x08, ...u16(no.length + (no.length > 0 ? 4 : 0) + ending.length - 1)];
                const lastGoto = no.length > 0 ? [0x19, 0x08, ...u16(ending.length - 1)] : [];
                return [
                    ...(encoded[0] ?? []),
                    0x19, 0x02, ...u16(yes.length + 4),
                    ...yes, ...firstGoto, ...no, ...lastGoto, ...ending,
                ];
            }
            if (info.id === 480) {
                const fallback = encoded[1] ?? [];
                return [
                    ...(encoded[0] ?? []),
                    0x19, 0x80, ...u16(fallback.length),
                    ...fallback,
                    0x19, 0x08, ...u16(ending.length - 1),
                    ...ending,
                ];
            }
            return [...encoded.flat(), ...ending];
        }
    }
}

function containsVolatile(node: Node): boolean {
    switch (node.kind) {
        case 'function':
            return VOLATILE.has(node.name) || node.args.some(containsVolatile);
        case 'paren':
        case 'percent':
        case 'unary':
            return containsVolatile(node.value);
        case 'binary':
            return containsVolatile(node.left) || containsVolatile(node.right);
        default:
            return false;
    }
}

/** Compile formula text (with or without a leading `=`) to BIFF12 tokens. */
export function compileXlsbFormula(formula: string): Uint8Array {
    const text = formula.trim().replace(/^=/, '');
    if (!text) throw new Error('FormulaCell.formula must not be empty');
    if (text.length > 8192) throw new RangeError('XLSB formula text is limited to 8192 characters');
    const tree = new Parser(tokenize(text)).parse();
    let tokens = emitNode(tree);
    if (containsVolatile(tree)) {
        tokens = [0x19, 0x01, 0x00, 0x00, ...tokens];
    }
    if (tokens.length === 0 || tokens.length > 16384) {
        throw new RangeError('XLSB formula token length exceeds the BIFF12 limit');
    }
    return Uint8Array.from(tokens);
}

function wideString(text: string): number[] {
    const out = [...u32(text.length)];
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        out.push(code & 0xff, (code >>> 8) & 0xff);
    }
    return out;
}

/**
 * Build a BrtFmla* record with the cached value Excel displays before recalc.
 * `cached` must already be converted from `Date` to an Excel serial number.
 */
export function buildXlsbFormulaRecord(
    column: number,
    styleIndex: number,
    formula: string,
    cached: string | number | boolean | ExcelError | null | undefined,
): Uint8Array {
    let recordId = BRT_FMLA_NUM;
    let cache: number[];
    if (cached instanceof ExcelError) {
        recordId = BRT_FMLA_ERROR;
        cache = [excelErrorToBiffCode(cached.code)];
    } else if (typeof cached === 'string') {
        recordId = BRT_FMLA_STRING;
        cache = wideString(cached);
    } else if (typeof cached === 'boolean') {
        recordId = BRT_FMLA_BOOL;
        cache = [cached ? 1 : 0];
    } else {
        const value = typeof cached === 'number' ? cached : 0;
        if (!Number.isFinite(value)) {
            throw new RangeError('XLSB formula cached values must be finite');
        }
        cache = doubleBytes(value);
    }

    const tokens = compileXlsbFormula(formula);
    const payload = [
        ...u32(column),
        ...u32(styleIndex),
        ...cache,
        0x01, 0x00,
        ...u32(tokens.length),
        ...tokens,
        ...u32(0),
    ];
    return buildBrtRecord(recordId, Uint8Array.from(payload));
}
