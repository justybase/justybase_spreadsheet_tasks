// Core classes
export { BigBuffer } from './BigBuffer';
export { ExcelReaderAbstract } from './ExcelReaderAbstract';
export { BiffReaderWriter } from './BiffReaderWriter';

// Readers
export { XlsbReader } from './XlsbReader';
export { XlsxReader } from './XlsxReader';
export { ReaderFactory } from './ReaderFactory';

// Writers
export { XlsbWriter } from './XlsbWriter';
export { XlsxWriter } from './XlsxWriter';

// Updater (in-place update of an existing workbook)
export { XlsxUpdater, XlsmUpdater } from './XlsxUpdater';
export type { ReplaceSheetDataOptions, RowSource } from './XlsxUpdater';
export { XlsbUpdater } from './XlsbUpdater';
export type { ReplaceSheetDataOptions as XlsbReplaceSheetDataOptions } from './XlsbUpdater';

// Excel COM conversion (Node.js + Windows + desktop Microsoft Excel)
export {
    convertXlsbToXlsx,
    convertXlsxToXlsb,
} from './ExcelConverter';
export type { ExcelConversionOptions } from './ExcelConverter';

// Formatting
export {
    FormattedCell,
    CellValue,
    PrimitiveCellValue,
    F,
    isFormattedCell,
    unwrapCell,
    getFormat,
} from './Formats';

// Writer helpers (shared)
export {
    sanitizeSheetName,
    initColWidths,
    applyHeaderWidths,
    updateColWidthsFromRows,
    defaultColWidth,
    INVALID_SHEET_NAME_CHARS,
    MAX_ROWS,
    MAX_COLS,
    MAX_CELL_CHARS,
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
} from './writerHelpers';
export { StreamingSheetState } from './StreamingSheetState';

// Shared XML helpers (tag attributes, rel targets, column letters)
export {
    escapeXmlText,
    unescapeXml,
    columnIndexToLetter,
    columnLetterToIndex,
    parseSharedStringsXml,
    trimTrailingEmptyRows,
    parseTagAttributes,
    resolveRelTarget,
} from './xmlUtils';
