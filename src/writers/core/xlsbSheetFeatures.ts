/**
 * XLSB worksheet feature records: merged cells, hyperlinks and data
 * validations. The XLSX-facing spec types are reused so callers pass the same
 * option objects to both writers.
 */
import { DataValidationSpec, HyperlinkSpec } from './xlsxParts';
import { buildBrtRecord, concat } from './xlsbRecords';
import { compileXlsbFormula } from './xlsbFormula';

const BRT_BEGIN_MERGE_CELLS = 0x00b1;
const BRT_MERGE_CELL = 0x00b0;
const BRT_END_MERGE_CELLS = 0x00b2;

const BRT_BEGIN_DVALS = 0x023d;
const BRT_DVAL = 0x0040;
const BRT_END_DVALS = 0x023e;

const BRT_HLINK = 0x01ee;

const DV_TYPES: Record<DataValidationSpec['type'], number> = {
    whole: 1,
    decimal: 2,
    list: 3,
    date: 4,
    time: 5,
    textLength: 6,
    custom: 7,
};

const DV_ERROR_STYLES: Record<string, number> = {
    stop: 0,
    warning: 1,
    information: 2,
};

const DV_OPERATORS: Record<string, number> = {
    between: 0,
    notBetween: 1,
    equal: 2,
    notEqual: 3,
    lessThan: 4,
    lessThanOrEqual: 5,
    greaterThan: 6,
    greaterThanOrEqual: 7,
};

interface CellRange {
    c1: number;
    r1: number;
    c2: number;
    r2: number;
}

function u16(value: number): number[] {
    const v = value & 0xffff;
    return [v & 0xff, (v >>> 8) & 0xff];
}

function u32(value: number): number[] {
    const v = value >>> 0;
    return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
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
    return text === undefined || text === '' ? [0xff, 0xff, 0xff, 0xff] : wideString(text);
}

function columnIndex(letters: string): number {
    let value = 0;
    for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
    return value - 1;
}

function parseCell(text: string): { col: number; row: number } {
    const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(text.trim());
    if (!match) throw new Error(`Invalid cell reference '${text}'`);
    const col = columnIndex(match[1]);
    const row = parseInt(match[2], 10) - 1;
    if (col < 0 || col > 16383 || row < 0 || row > 1048575) {
        throw new RangeError(`Cell reference '${text}' is outside the Excel grid`);
    }
    return { col, row };
}

/** Parse `A1:B2` (or a single `A1`) into zero-based coordinates. */
export function parseXlsbRange(text: string): CellRange {
    const parts = text.split(':');
    if (parts.length > 2) throw new Error(`Invalid range '${text}'`);
    const first = parseCell(parts[0]);
    const last = parts.length === 2 ? parseCell(parts[1]) : first;
    return {
        c1: Math.min(first.col, last.col),
        r1: Math.min(first.row, last.row),
        c2: Math.max(first.col, last.col),
        r2: Math.max(first.row, last.row),
    };
}

/** BrtBeginMergeCells / BrtMergeCell* / BrtEndMergeCells. */
export function buildXlsbMergeRecords(merges: string[]): Uint8Array {
    if (merges.length === 0) return new Uint8Array(0);
    const parts: Uint8Array[] = [
        buildBrtRecord(BRT_BEGIN_MERGE_CELLS, Uint8Array.from(u32(merges.length))),
    ];
    for (const ref of merges) {
        const range = parseXlsbRange(ref);
        parts.push(buildBrtRecord(BRT_MERGE_CELL, Uint8Array.from([
            ...u32(range.r1), ...u32(range.r2), ...u32(range.c1), ...u32(range.c2),
        ])));
    }
    parts.push(buildBrtRecord(BRT_END_MERGE_CELLS, new Uint8Array(0)));
    return concat(parts);
}

/** One BrtHLink record. `rId` is required for external targets. */
export function buildXlsbHyperlinkRecord(
    row: number,
    col: number,
    link: HyperlinkSpec,
    rId?: string,
): Uint8Array {
    const internal = link.target === undefined;
    const relationship = internal ? '' : (rId ?? '');
    const location = internal ? (link.location ?? '') : '';
    const display = link.display
      ?? (internal ? (link.location ?? '') : (link.target ?? ''));
    const payload = [
        ...u32(row), ...u32(row), ...u32(col), ...u32(col),
        ...wideString(relationship),
        ...wideString(location),
        ...wideString(link.tooltip ?? ''),
        ...wideString(display),
    ];
    return buildBrtRecord(BRT_HLINK, Uint8Array.from(payload));
}

function literalListFormula(values: string[]): string {
    return `"${values.map(value => value.replace(/"/g, '""')).join(',')}"`;
}

function validationFormula(text: string | undefined): Uint8Array {
    if (text === undefined || text === '') return new Uint8Array(8);
    const tokens = compileXlsbFormula(text);
    return Uint8Array.from([...u32(tokens.length), ...tokens, ...u32(0)]);
}

/** BrtBeginDVals / BrtDVal* / BrtEndDVals. */
export function buildXlsbValidationRecords(validations: DataValidationSpec[]): Uint8Array {
    if (validations.length === 0) return new Uint8Array(0);
    const parts: Uint8Array[] = [
        buildBrtRecord(BRT_BEGIN_DVALS, Uint8Array.from([
            ...u16(0), ...u32(0), ...u32(0), ...u32(0), ...u32(validations.length),
        ])),
    ];
    for (const rule of validations) {
        const ranges = rule.sqref.trim().split(/\s+/).map(parseXlsbRange);
        if (ranges.length === 0) throw new Error('Data validation sqref must not be empty');

        let flags = DV_TYPES[rule.type] ?? 0;
        flags |= (DV_ERROR_STYLES[rule.errorStyle ?? 'stop'] ?? 0) << 4;
        const isLiteralList = rule.type === 'list' && rule.values !== undefined;
        if (isLiteralList) flags |= 0x80;
        if (rule.allowBlank) flags |= 0x100;
        if (rule.showDropdown === false) flags |= 0x200;
        if (rule.showInputMessage) flags |= 1 << 18;
        if (rule.showErrorMessage) flags |= 1 << 19;
        if (rule.operator !== undefined) flags |= (DV_OPERATORS[rule.operator] ?? 0) << 20;

        const payload: number[] = [...u32(flags), ...u32(ranges.length)];
        for (const range of ranges) {
            payload.push(...u32(range.r1), ...u32(range.r2), ...u32(range.c1), ...u32(range.c2));
        }
        payload.push(
            ...nullableString(rule.errorTitle),
            ...nullableString(rule.error),
            ...nullableString(rule.promptTitle),
            ...nullableString(rule.prompt),
        );
        const formula1 = isLiteralList
            ? literalListFormula(rule.values!)
            : rule.formula1;
        payload.push(...validationFormula(formula1));
        payload.push(...validationFormula(rule.formula2));
        parts.push(buildBrtRecord(BRT_DVAL, Uint8Array.from(payload)));
    }
    parts.push(buildBrtRecord(BRT_END_DVALS, new Uint8Array(0)));
    return concat(parts);
}
