import { CellValue } from './Formats';
import { oaDateToTime } from './writerHelpers';

/**
 * Shared reader surface for XLSB and XLSX.
 * Always use `await reader.read()` — both formats are async in v2+.
 */
export abstract class ExcelReaderAbstract {
    fieldCount: number = 0;
    rowCount: number = 0;
    actualSheetName: string = '';
    resultsCount: number = 0;

    protected _oaEpoch: number;

    constructor() {
        this._oaEpoch = Date.UTC(1899, 11, 30);
    }

    /** Open a workbook file. */
    abstract open(path: string, readSharedStrings?: boolean, updateMode?: boolean): Promise<void>;

    async close(): Promise<void> {
        // cleanup if needed
    }

    /**
     * Advance to the next row.
     * @returns true if a row is available; false at EOF
     */
    abstract read(): Promise<boolean>;

    /** Worksheet names in workbook order. */
    abstract getSheetNames(): string[];

    /** Cell value for column index `i` of the current row (0-based). */
    abstract getValue(i: number): CellValue;

    /**
     * Select a worksheet for subsequent `read()` calls (by name or 0-based index).
     * Resets row state/EOF; default is sheet 0. Resets `fieldCount`.
     */
    abstract selectSheet(nameOrIndex: string | number): Promise<void>;

    /**
     * Advance to the next worksheet (like `nextResult` in DB readers).
     * @returns true if another sheet was selected; false at end.
     */
    abstract nextResult(): Promise<boolean>;

    dispose(): void {
        // cleanup
    }

    getDateTimeFromOaDate(oaDate: number): Date {
        // Round to the nearest millisecond: the float product can land just below a whole ms (x.999),
        // and Date truncates instead of rounding. Handles the fictitious 1900-02-29.
        return new Date(oaDateToTime(oaDate, this._oaEpoch));
    }
}

export default ExcelReaderAbstract;
