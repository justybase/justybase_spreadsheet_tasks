/**
 * Browser Spreadsheet - Entry point for browser-based Excel file generation.
 * Features: autofit columns, frozen+bold header row, auto-filter.
 */
import { BrowserXlsxWriter } from './BrowserXlsxWriter';
import { BrowserXlsbWriter } from './BrowserXlsbWriter';
import { CellValue } from '../Formats';

export { BrowserXlsxWriter } from './BrowserXlsxWriter';
export type { BrowserSheetOptions, BrowserStartSheetOptions } from './BrowserXlsxWriter';
export { BrowserXlsbWriter } from './BrowserXlsbWriter';
export type {
    BrowserSheetOptions as BrowserXlsbSheetOptions,
    BrowserStartSheetOptions as BrowserXlsbStartSheetOptions,
} from './BrowserXlsbWriter';
export { BrowserBigBuffer } from './BrowserBigBuffer';
export { BrowserZip } from './BrowserZip';
export { F, type CellValue, type FormattedCell, type CellStyle, type StyledCell, type FormulaCell, type RichText, type RichTextRun } from '../Formats';

/** Trigger browser download of a Blob */
function triggerDownload(blob: Blob, fileName: string): void {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}

export interface BrowserSheetInput {
    name: string;
    rows: CellValue[][];
    headers: string[];
}

/** Download data as XLSX. Headers = bold + frozen + auto-filter. Columns = autofit. */
export function downloadXlsx(
    fileName: string,
    rows: CellValue[][],
    headers: string[],
    sheetName: string = 'Sheet1'
): void {
    const w = new BrowserXlsxWriter();
    w.addSheet(sheetName);
    w.writeSheet(rows, headers, { doAutofilter: true });
    triggerDownload(w.finalize(), fileName);
}

/** Download multiple sheets as XLSX. */
export function downloadXlsxMultiSheet(fileName: string, sheets: BrowserSheetInput[]): void {
    const w = new BrowserXlsxWriter();
    for (const s of sheets) {
        w.addSheet(s.name);
        w.writeSheet(s.rows, s.headers, { doAutofilter: true });
    }
    triggerDownload(w.finalize(), fileName);
}

/** Download data as XLSB. Headers = bold + frozen + auto-filter. Columns = autofit. */
export function downloadXlsb(
    fileName: string,
    rows: CellValue[][],
    headers: string[],
    sheetName: string = 'Sheet1'
): void {
    const w = new BrowserXlsbWriter();
    w.addSheet(sheetName);
    w.writeSheet(rows, headers, { doAutofilter: true });
    triggerDownload(w.finalize(), fileName);
}

/** Download multiple sheets as XLSB. */
export function downloadXlsbMultiSheet(fileName: string, sheets: BrowserSheetInput[]): void {
    const w = new BrowserXlsbWriter();
    for (const s of sheets) {
        w.addSheet(s.name);
        w.writeSheet(s.rows, s.headers, { doAutofilter: true });
    }
    triggerDownload(w.finalize(), fileName);
}
