import { isFormattedCell } from './Formats';

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

/** Apply autofit widths from header labels. */
export function applyHeaderWidths(colWidths: number[], headers: string[], columnCount: number): void {
    for (let i = 0; i < columnCount; i++) {
        const len = headers[i] ? headers[i].length + 1 : 0;
        let width = 1.25 * len + 2;
        if (width > 80) width = 80;
        if (colWidths[i] < width) colWidths[i] = width;
    }
}

/**
 * Sample up to `sampleLimit` rows and expand column widths (autofit).
 * Skips null/undefined cells; Dates use a fixed display length of 20.
 */
export function updateColWidthsFromRows(
    colWidths: number[],
    rows: unknown[][],
    sampleLimit: number = 100
): void {
    for (let r = 0; r < Math.min(rows.length, sampleLimit); r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
            const raw = row[c];
            if (raw === null || raw === undefined) continue;
            const val = isFormattedCell(raw) ? raw.value : raw;
            const len = val instanceof Date ? 20 : String(val).length + 1;
            let width = 1.25 * len + 2;
            if (width > 80) width = 80;
            if (colWidths[c] < width) colWidths[c] = width;
        }
    }
}

/**
 * Resolve a stored autofit width to the value written into the file.
 * XLSB uses floored integers; XLSX keeps the float.
 */
export function defaultColWidth(width: number, floor: boolean = false): number {
    if (width > 0) {
        return floor ? Math.floor(width) : width;
    }
    return 10;
}
