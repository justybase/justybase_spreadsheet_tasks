/**
 * BIFF12 sheet-layout records shared by the Node and browser XLSB writers:
 * panes, selections and column records.
 */
import { StyleColor } from '../../Formats';
import {
    FreezePane,
    Outline,
    SheetProtectionSpec,
    colorKey,
    passwordHash,
} from './sheetLayout';
import { buildBrtRecord, concat } from './xlsbRecords';

export const BRT_PANE = 0x0097;
export const BRT_SEL = 0x0098;
export const BRT_BEGIN_COL_INFOS = 0x0186;
export const BRT_COL_INFO = 0x003c;
export const BRT_END_COL_INFOS = 0x0187;

const PANE_ACTIVE: Record<FreezePane['activePane'], number> = {
    bottomRight: 0,
    topRight: 1,
    bottomLeft: 2,
    topLeft: 3,
};

/** 8-byte BrtColor payload; all-zero bytes mean automatic. */
export function xlsbColorPayload(color: StyleColor | undefined): Uint8Array {
    if (color === undefined) return new Uint8Array(8);
    const key = colorKey(color);
    if (key === 'auto') return Uint8Array.from([0x01, 0x40, 0, 0, 0, 0, 0, 0]);
    if (key.startsWith('rgb:')) {
        const value = key.slice(4);
        const alpha = parseInt(value.slice(0, 2), 16);
        const red = parseInt(value.slice(2, 4), 16);
        const green = parseInt(value.slice(4, 6), 16);
        const blue = parseInt(value.slice(6, 8), 16);
        return Uint8Array.from([0x05, 0, 0, 0, red, green, blue, alpha]);
    }
    if (key.startsWith('indexed:')) {
        const index = parseInt(key.slice(8), 10) & 0xff;
        return Uint8Array.from([0x03, index, 0, 0, 0, 0, 0, 0xff]);
    }
    const [, themeText, tintText] = key.split(':');
    const theme = parseInt(themeText, 10) & 0xff;
    const tint = Math.max(-32768, Math.min(32767, Math.round(Number(tintText) * 32767)));
    return Uint8Array.from([
        0x07, theme, tint & 0xff, (tint >> 8) & 0xff, 0, 0, 0, 0xff,
    ]);
}

/** BrtPane ([MS-XLSB] 2.4.716). */
export function buildXlsbPaneRecord(freeze: FreezePane | null): Uint8Array {
    const xSplit = freeze?.xSplit ?? 0;
    const ySplit = freeze?.ySplit ?? 0;
    const active = PANE_ACTIVE[freeze?.activePane ?? 'topLeft'];
    const flags = freeze ? 0x03 : 0x00;
    const payload = new Uint8Array(8 + 8 + 4 + 4 + 4 + 1);
    const view = new DataView(payload.buffer);
    view.setFloat64(0, xSplit, true);
    view.setFloat64(8, ySplit, true);
    view.setUint32(16, ySplit, true);
    view.setUint32(20, xSplit, true);
    view.setUint32(24, active, true);
    payload[28] = flags;
    return buildBrtRecord(BRT_PANE, payload);
}

/** BrtSel ([MS-XLSB] 2.4.781). */
export function buildXlsbSelectionRecord(freeze: FreezePane | null): Uint8Array {
    const active = PANE_ACTIVE[freeze?.activePane ?? 'topLeft'];
    const row = freeze?.ySplit ?? 0;
    const col = freeze?.xSplit ?? 0;
    const payload = new Uint8Array(20 + 16);
    const view = new DataView(payload.buffer);
    view.setUint32(0, active, true);
    view.setUint32(4, row, true);
    view.setUint32(8, col, true);
    view.setUint32(12, 0, true);
    view.setUint32(16, 1, true);
    view.setUint32(20, row, true);
    view.setUint32(24, row, true);
    view.setUint32(28, col, true);
    view.setUint32(32, col, true);
    return buildBrtRecord(BRT_SEL, payload);
}

export interface XlsbColumnLayout {
    startCol: number;
    endCol: number;
    widths: number[];
    hiddenColumns: number[];
    outline?: Outline;
}

/** BrtBeginColInfos / BrtColInfo* / BrtEndColInfos. */
const BRT_SHEET_PROTECTION = 0x0217;
const BRT_LEGACY_DRAWING = 0x0227;
/** The BrtSheetProtection record starts at this offset in the sheet template. */
export const XLSB_PROTECTION_TEMPLATE_OFFSET = 221;

/** BrtSheetProtection ([MS-XLSB] 2.4.778): hash + 16 protection flags. */
export function buildXlsbSheetProtectionRecord(spec: SheetProtectionSpec): Uint8Array {
    // BrtSheetProtection stores "allowed" flags, the inverse of the OOXML
    // lock attributes. Selection flags already mean "allow selection".
    // XLSX omits the attribute to keep the default (blocked); only an explicit
    // false allows the action. Mirror that default here.
    const allowed = (blocked: boolean | undefined): number => (blocked === false ? 1 : 0);
    const selection = (allow: boolean | undefined): number => (allow === false ? 0 : 1);
    // Order follows the BrtSheetProtection bitfield: locked, objects, scenarios,
    // format cells/columns/rows, insert columns/rows/hyperlinks, delete
    // columns/rows, select locked, sort, autoFilter, pivot tables, select unlocked.
    const flags = [
        1,
        spec.objects === false ? 1 : 0,
        spec.scenarios === false ? 1 : 0,
        allowed(spec.formatCells),
        allowed(spec.formatColumns),
        allowed(spec.formatRows),
        allowed(spec.insertColumns),
        allowed(spec.insertRows),
        allowed(spec.insertHyperlinks),
        allowed(spec.deleteColumns),
        allowed(spec.deleteRows),
        selection(spec.selectLockedCells),
        allowed(spec.sort),
        allowed(spec.autoFilter),
        allowed(spec.pivotTables),
        selection(spec.selectUnlockedCells),
    ];
    const payload = new Uint8Array(2 + flags.length * 4);
    const view = new DataView(payload.buffer);
    view.setUint16(0, spec.password ? passwordHash(spec.password) : 0, true);
    flags.forEach((value, index) => view.setUint32(2 + index * 4, value, true));
    return buildBrtRecord(BRT_SHEET_PROTECTION, payload);
}

/**
 * Replace the template's unprotected BrtSheetProtection record when the sheet
 * is protected. The template already contains a 66-byte payload record at a
 * fixed offset, so no other offsets move.
 */
export function applyXlsbSheetProtection(
    template: Uint8Array,
    spec: SheetProtectionSpec | undefined,
): void {
    if (!spec || spec.enabled === false) return;
    const record = buildXlsbSheetProtectionRecord(spec);
    if (XLSB_PROTECTION_TEMPLATE_OFFSET + record.length > template.length) {
        throw new Error('XLSB sheet template is too small for the protection record');
    }
    template.set(record, XLSB_PROTECTION_TEMPLATE_OFFSET);
}

/** BrtLegacyDrawing: the relationship id of the comment VML drawing. */
export function buildXlsbLegacyDrawingRecord(relId: string): Uint8Array {
    const payload = new Uint8Array(4 + relId.length * 2);
    const view = new DataView(payload.buffer);
    view.setUint32(0, relId.length, true);
    for (let index = 0; index < relId.length; index++) {
        view.setUint16(4 + index * 2, relId.charCodeAt(index), true);
    }
    return buildBrtRecord(BRT_LEGACY_DRAWING, payload);
}

export function buildXlsbColumnRecords(layout: XlsbColumnLayout): Uint8Array {
    const hidden = new Set(layout.hiddenColumns);
    let maxCol = Math.max(layout.endCol - 1, layout.startCol);
    for (const col of hidden) maxCol = Math.max(maxCol, col);
    if (layout.outline) for (const col of layout.outline.indices) maxCol = Math.max(maxCol, col);

    const parts: Uint8Array[] = [buildBrtRecord(BRT_BEGIN_COL_INFOS, new Uint8Array(0))];
    for (let col = layout.startCol; col <= maxCol; col++) {
        const width = layout.widths[col] ?? 8.43;
        let flags = 0x02;
        if (hidden.has(col)) flags |= 0x01;
        if (layout.outline && col < layout.outline.level.length) {
            flags |= layout.outline.level[col] << 8;
            if (layout.outline.collapsed[col]) flags |= 0x1000;
        }
        const payload = new Uint8Array(18);
        const view = new DataView(payload.buffer);
        view.setUint32(0, col, true);
        view.setUint32(4, col, true);
        view.setUint32(8, Math.max(0, Math.min(255 * 256, Math.round(width * 256))), true);
        view.setUint32(12, 0, true);
        view.setUint16(16, flags, true);
        parts.push(buildBrtRecord(BRT_COL_INFO, payload));
    }
    parts.push(buildBrtRecord(BRT_END_COL_INFOS, new Uint8Array(0)));
    return concat(parts);
}
