/**
 * XLSB cell notes: the `xl/comments/commentN.bin` records and the classic VML
 * drawing that positions the note boxes. The VML is shared with the XLSX
 * writer so both formats display notes identically.
 */
import { CommentSpec, buildXlsxCommentsVml } from './xlsxParts';
import { buildBrtRecord, concat } from './xlsbRecords';
import { parseCellRef } from './sheetLayout';

const BEGIN_COMMENTS = 0x274;
const END_COMMENTS = 0x275;
const BEGIN_AUTHORS = 0x276;
const END_AUTHORS = 0x277;
const AUTHOR = 0x278;
const BEGIN_LIST = 0x279;
const END_LIST = 0x27a;
const BEGIN_COMMENT = 0x27b;
const END_COMMENT = 0x27c;
const TEXT = 0x27d;

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

/** Stable GUID derived from arbitrary text (comments, tables, ACUID blocks). */
export function deterministicGuid(text: string): Uint8Array {
    let hash = 0x811c9dc5;
    const bytes = new Uint8Array(16);
    for (let block = 0; block < 4; block++) {
        for (let index = 0; index < text.length; index++) {
            hash ^= text.charCodeAt(index) + block * 31;
            hash = Math.imul(hash, 0x01000193) >>> 0;
        }
        bytes[block * 4] = hash & 0xff;
        bytes[block * 4 + 1] = (hash >>> 8) & 0xff;
        bytes[block * 4 + 2] = (hash >>> 16) & 0xff;
        bytes[block * 4 + 3] = (hash >>> 24) & 0xff;
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x50; // version 5
    bytes[8] = (bytes[8] & 0x3f) | 0x80; // RFC 4122 variant
    return bytes;
}

/** Stable comment GUID derived from the part name and cell reference. */
function commentGuid(partName: string, ref: string): Uint8Array {
    return deterministicGuid(`${partName}:${ref}`);
}

/** `xl/comments/commentN.bin`. */
export function buildXlsbCommentsPart(comments: CommentSpec[], partName: string): Uint8Array {
    if (comments.length === 0) return new Uint8Array(0);
    const authors: string[] = [];
    const authorIds = new Map<string, number>();
    for (const comment of comments) {
        const author = comment.author ?? 'Author';
        if (!authorIds.has(author)) {
            authorIds.set(author, authors.length);
            authors.push(author);
        }
    }

    const parts: Uint8Array[] = [
        buildBrtRecord(BEGIN_COMMENTS, new Uint8Array(0)),
        buildBrtRecord(BEGIN_AUTHORS, new Uint8Array(0)),
    ];
    for (const author of authors) {
        parts.push(buildBrtRecord(AUTHOR, Uint8Array.from(wideString(author))));
    }
    parts.push(buildBrtRecord(END_AUTHORS, new Uint8Array(0)));
    parts.push(buildBrtRecord(BEGIN_LIST, new Uint8Array(0)));

    for (const comment of comments) {
        const { row, col } = parseCellRef(comment.ref);
        const authorId = authorIds.get(comment.author ?? 'Author')!;
        const guid = commentGuid(partName, comment.ref);
        const header = Uint8Array.from([
            ...u32(authorId), ...u32(row), ...u32(row), ...u32(col), ...u32(col),
            ...guid,
        ]);
        parts.push(buildBrtRecord(BEGIN_COMMENT, header));
        // RichStr: fRichStr = 0, UTF-16 text without formatting runs.
        parts.push(buildBrtRecord(TEXT, Uint8Array.from([0, ...wideString(comment.text)])));
        parts.push(buildBrtRecord(END_COMMENT, new Uint8Array(0)));
    }

    parts.push(buildBrtRecord(END_LIST, new Uint8Array(0)));
    parts.push(buildBrtRecord(END_COMMENTS, new Uint8Array(0)));
    return concat(parts);
}

/** The VML drawing part that positions the note shapes. */
export function buildXlsbCommentsVml(comments: CommentSpec[]): string {
    return buildXlsxCommentsVml(comments);
}
