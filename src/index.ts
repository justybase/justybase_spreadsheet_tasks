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
export type { XlsbSheetOptions } from './XlsbWriter';
export { XlsxWriter } from './XlsxWriter';
export type { XlsxSheetOptions } from './XlsxWriter';
export type {
    XlsxSheetFeatures,
    HyperlinkSpec,
    DataValidationSpec,
    ConditionalFormatSpec,
    ConditionalFormatRule,
    CommentSpec,
} from './writers/core/xlsxParts';

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
    ExcelError,
    RichText,
    RichTextRun,
    CellStyle,
    CellValue,
    FontStyle,
    FontUnderline,
    FillStyle,
    PatternType,
    BorderStyle,
    BorderSide,
    BorderLineStyle,
    AlignmentStyle,
    ProtectionStyle,
    StyledCell,
    FormulaCell,
    FormattedCell,
    PrimitiveCellValue,
    StyleColor,
    F,
    isFormattedCell,
    isStyledCell,
    isFormulaCell,
    isRichText,
    getCellStyle,
    getFormula,
    mergeCellStyle,
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
    formatColumnWidth,
    assertColumnWidth,
    autofitWidth,
    autofitFileWidth,
    textDisplayUnits,
    DEFAULT_COLUMN_WIDTH,
    MAX_AUTOFIT_WIDTH,
    AUTOFIT_SAMPLE_ROWS,
    resolveAutofitRows,
    COLUMN_WIDTH_PADDING,
    AUTOFILTER_BUTTON_WIDTH,
    INVALID_SHEET_NAME_CHARS,
    MAX_ROWS,
    MAX_COLS,
    MAX_CELL_CHARS,
    assertCellTextLength,
    assertColumnCount,
    assertRowFitsColumns,
    assertRowCount,
} from './writerHelpers';
export type { HeaderWidthOptions } from './writerHelpers';
export { StreamingSheetState } from './StreamingSheetState';

// Shared XML helpers (tag attributes, rel targets, column letters)
export {
    escapeXmlText,
    escapeXlsxText,
    unescapeXml,
    unescapeXmlText,
    columnIndexToLetter,
    columnLetterToIndex,
    parseSharedStringsXml,
    parseSharedStringsBuffer,
    trimTrailingEmptyRows,
    parseTagAttributes,
    resolveRelTarget,
} from './xmlUtils';
