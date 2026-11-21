/**
 * XLSB table parts (`xl/tables/tableN.bin`) and the worksheet's
 * BrtBeginListParts block. The record layout follows the files Excel writes
 * for the same table saved from an XLSX workbook.
 */
import { NormalizedTableSpec, tableAutoFilterRef } from './sheetLayout';
import { buildBrtRecord, concat } from './xlsbRecords';
import { compileXlsbFormulaStats } from './xlsbFormula';
import { deterministicGuid } from './xlsbComments';

const BRT_AC_BEGIN = 0x25;
const BRT_FRT_BEGIN = 0xc00;
const BRT_AC_END = 0x26;
const BRT_BEGIN_LIST = 0x157;
const BRT_END_LIST = 0x158;
const BRT_BEGIN_LIST_COLS = 0x159;
const BRT_END_LIST_COLS = 0x15a;
const BRT_BEGIN_LIST_COL = 0x15b;
const BRT_END_LIST_COL = 0x15c;
const BRT_LIST_TR_FMLA = 0x160;
const BRT_TABLE_STYLE_CLIENT = 0x201;
const BRT_BEGIN_AUTOFILTER = 0xa1;
const BRT_END_AUTOFILTER = 0xa2;
const BRT_BEGIN_LIST_PARTS = 0x294;
const BRT_LIST_PART = 0x295;
const BRT_END_LIST_PARTS = 0x296;

const TOTALS_CODES: Record<string, number> = {
    average: 1,
    count: 2,
    countNums: 3,
    max: 4,
    min: 5,
    sum: 6,
    stdDev: 7,
    var: 8,
    custom: 9,
};

function u32(value: number): number[] {
    const v = value >>> 0;
    return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
}

function i32(value: number): number[] {
    return u32(value >>> 0);
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

function rangeBytes(range: { c1: number; r1: number; c2: number; r2: number }): number[] {
    return [...u32(range.r1), ...u32(range.r2), ...u32(range.c1), ...u32(range.c2)];
}

/** BrtACBegin / BrtFRTBegin (revision GUID) / BrtACEnd. */
function acuid(key: string): Uint8Array {
    return concat([
        buildBrtRecord(BRT_AC_BEGIN, Uint8Array.from([1, 0, 2, 0x10, 0, 0x80])),
        buildBrtRecord(BRT_FRT_BEGIN, deterministicGuid(key)),
        buildBrtRecord(BRT_AC_END, new Uint8Array(0)),
    ]);
}

/** Build one `xl/tables/tableN.bin` part. */
export function buildXlsbTablePart(table: NormalizedTableSpec, tableId: number): Uint8Array {
    const parts: Uint8Array[] = [acuid(`${tableId}:${table.name}`)];

    const beginList: number[] = [
        ...rangeBytes(table.ref),
        ...u32(0), // lt = LTRANGE
        ...u32(tableId),
        ...u32(table.headerRow ? 1 : 0),
        ...u32(table.totalsRow ? 1 : 0),
        ...u32(table.totalsRow ? 1 : 0), // fShowTotals
    ];
    for (let index = 0; index < 6; index++) beginList.push(...i32(-1));
    beginList.push(...u32(0)); // dwConnID
    beginList.push(...nullableString(table.name));
    beginList.push(...nullableString(table.name));
    beginList.push(...u32(0)); // stComment: empty string
    for (let index = 0; index < 3; index++) beginList.push(...nullableString(undefined));
    parts.push(buildBrtRecord(BRT_BEGIN_LIST, Uint8Array.from(beginList)));
    parts.push(acuid(`${tableId}:${table.name}:list`));

    if (table.headerRow) {
        const filterRef = tableAutoFilterRef(table);
        if (filterRef !== null) {
            const endRow = table.totalsRow ? table.ref.r2 - 1 : table.ref.r2;
            parts.push(buildBrtRecord(BRT_BEGIN_AUTOFILTER, Uint8Array.from([
                ...u32(table.ref.r1), ...u32(endRow), ...u32(table.ref.c1), ...u32(table.ref.c2),
            ])));
            parts.push(buildBrtRecord(BRT_END_AUTOFILTER, new Uint8Array(0)));
        }
    }

    parts.push(buildBrtRecord(BRT_BEGIN_LIST_COLS, Uint8Array.from(u32(table.columns.length))));
    table.columns.forEach((column, index) => {
        const number = index + 1;
        parts.push(acuid(`${tableId}:${table.name}:${number}`));
        const code = column.totalsRowFunction === undefined
            ? 0
            : TOTALS_CODES[column.totalsRowFunction] ?? 0;
        const columnPayload: number[] = [
            ...u32(number),
            ...u32(code),
            ...i32(-1), ...i32(-1), ...i32(-1),
            ...u32(0), // idqsif
            ...nullableString(undefined),
            ...nullableString(column.name),
            ...nullableString(column.totalsRowLabel),
            ...nullableString(undefined),
            ...nullableString(undefined),
            ...nullableString(undefined),
        ];
        parts.push(buildBrtRecord(BRT_BEGIN_LIST_COL, Uint8Array.from(columnPayload)));
        if (code === TOTALS_CODES.custom && column.totalsRowFormula !== undefined) {
            const compiled = compileXlsbFormulaStats(column.totalsRowFormula);
            parts.push(buildBrtRecord(BRT_LIST_TR_FMLA, Uint8Array.from([
                0, ...u32(compiled.tokens.length), ...compiled.tokens, ...u32(0),
            ])));
        }
        parts.push(buildBrtRecord(BRT_END_LIST_COL, new Uint8Array(0)));
    });
    parts.push(buildBrtRecord(BRT_END_LIST_COLS, new Uint8Array(0)));

    let flags = 0;
    if (table.showFirstColumn) flags |= 0x01;
    if (table.showLastColumn) flags |= 0x02;
    if (table.showRowStripes) flags |= 0x04;
    if (table.showColumnStripes) flags |= 0x08;
    parts.push(buildBrtRecord(BRT_TABLE_STYLE_CLIENT, Uint8Array.from([
        flags & 0xff, (flags >>> 8) & 0xff,
        ...nullableString(table.styleName),
    ])));
    parts.push(buildBrtRecord(BRT_END_LIST, new Uint8Array(0)));
    return concat(parts);
}

/** BrtBeginListParts / BrtListPart* / BrtEndListParts for one worksheet. */
export function buildXlsbListParts(relIds: string[]): Uint8Array {
    if (relIds.length === 0) return new Uint8Array(0);
    const parts: Uint8Array[] = [
        buildBrtRecord(BRT_BEGIN_LIST_PARTS, Uint8Array.from(u32(relIds.length))),
    ];
    for (const relId of relIds) {
        parts.push(buildBrtRecord(BRT_LIST_PART, Uint8Array.from(wideString(relId))));
    }
    parts.push(buildBrtRecord(BRT_END_LIST_PARTS, new Uint8Array(0)));
    return concat(parts);
}
