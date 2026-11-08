/**
 * Shared low-level XML helpers used by readers and the XlsxUpdater.
 */

/** True when the string contains XML-invalid control characters (tab/CR/LF are allowed). */
function hasInvalidControlChars(str: string): boolean {
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        if (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
            return true;
        }
    }
    return false;
}

/** Remove XML-invalid control characters (tab/CR/LF and normal ranges are kept). */
function stripInvalidControlChars(str: string): string {
    let out = '';
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        if (c === 0x09 || c === 0x0a || c === 0x0d || (c >= 0x20 && c <= 0xd7ff) || (c >= 0xe000 && c <= 0xfffd)) {
            out += str[i];
        }
    }
    return out;
}

/** Escape text content for use inside an XML element body. */
export function escapeXmlText(str: string): string {
    if (!/[&<>"'\r]/.test(str) && !hasInvalidControlChars(str)) {
        return str;
    }
    // CR is written as &#13; so readers that normalize raw line ends keep it.
    let out = str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;')
        .replace(/\r/g, '&#13;');
    if (hasInvalidControlChars(out)) {
        out = stripInvalidControlChars(out);
    }
    return out;
}

/** Unescape the five predefined XML entities (single-pass, order-independent). */
export function unescapeXml(str: string): string {
    if (!str || str.indexOf('&') === -1) return str;
    return str.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (entity: string) => {
        switch (entity) {
            case '&amp;': return '&';
            case '&lt;': return '<';
            case '&gt;': return '>';
            case '&quot;': return '"';
            case '&apos;': return "'";
            default: {
                // Numeric character reference: &#123; or &#x1F600;
                if (entity[1] === '#') {
                    const body = entity.slice(2, -1);
                    const code = body[0] === 'x' || body[0] === 'X'
                        ? parseInt(body.slice(1), 16)
                        : parseInt(body, 10);
                    if (Number.isFinite(code)) {
                        try { return String.fromCodePoint(code); } catch { return entity; }
                    }
                }
                return entity;
            }
        }
    });
}

/**
 * Escape text of a shared string or inline string. Beyond XML escaping, a literal
 * `_xHHHH_` must be written as `_x005F_xHHHH_` because Excel decodes that pattern
 * in these elements (ECMA-376 ST_Xstring). Attribute values use `escapeXmlText`.
 */
export function escapeXlsxText(str: string): string {
    return escapeXmlText(str.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_'));
}

/**
 * Unescape element text content. Raw CR LF and lone CR are normalized to LF first
 * (XML 1.0 §2.11); character references such as `&#13;` are kept as CR. Then the
 * ST_Xstring escape `_xHHHH_` that Excel writes for control characters is decoded.
 */
export function unescapeXmlText(str: string): string {
    if (!str) return str;
    const text = unescapeXml(str.indexOf('\r') === -1 ? str : str.replace(/\r\n?/g, '\n'));
    return text.indexOf('_x') === -1
        ? text
        : text.replace(/_x([0-9A-Fa-f]{4})_/g, (_match: string, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

/** Parse XML tag attributes into a map, independent of order/prefix/quote style. */
export function parseTagAttributes(tag: string): Record<string, string> {
    const attrs: Record<string, string> = {};
    // Matches name="value" or name='value', allowing : and . in names (r:id, xml:space).
    const re = /([A-Za-z_][\w.:-]*)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(tag)) !== null) {
        const name = m[1];
        const value = m[3] !== undefined ? m[3] : m[4];
        attrs[name] = value;
        // Also store local name without prefix for convenience (r:id -> id).
        const colon = name.indexOf(':');
        if (colon !== -1) {
            const local = name.slice(colon + 1);
            if (!(local in attrs)) attrs[local] = value;
        }
    }
    return attrs;
}

/** Resolve a workbook rel Target to a zip entry path (xlspy _resolve_path logic). */
export function resolveRelTarget(target: string): string {
    if (!target) return target;
    if (target.startsWith('/')) return target.slice(1);
    if (target.startsWith('xl/')) return target;
    return 'xl/' + target;
}

/** 0-based column index -> "A", "B", ..., "AA", "AB", ... (bijective base-26). */
export function columnIndexToLetter(colIndex: number): string {
    let n = colIndex + 1;
    let result = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        result = String.fromCharCode(65 + rem) + result;
        n = Math.floor((n - 1) / 26);
    }
    return result;
}

/** "A", "AB", ... -> 0-based column index. */
export function columnLetterToIndex(letter: string): number {
    let column = 0;
    const length = letter.length;
    for (let i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column - 1;
}

/**
 * Remove trailing rows that contain no non-empty cells (padding rows sometimes
 * produced by SQL exports). This keeps the written range aligned with the
 * actual data, so pivot tables do not pick up "(blank)" items from empty rows.
 */
export function trimTrailingEmptyRows<T>(rows: T[][]): T[][] {
    let end = rows.length;
    while (end > 0 && rows[end - 1].every(v => v === null || v === undefined)) {
        end--;
    }
    return end === rows.length ? rows : rows.slice(0, end);
}

/**
 * Parse the contents of xl/sharedStrings.xml into the list of string values.
 * Handles plain `<si><t>…</t></si>` entries as well as rich-text `<si><r><t>…</t></r></si>` runs.
 * Phonetic `<rPh>` blocks are skipped (xlspy parity) and self-closing `<si/>` yields "".
 */
export function parseSharedStringsXml(xml: string): string[] {
    const result: string[] = [];
    // Self-closing <si/> entries (empty strings) — count them to keep indices aligned.
    // Full <si>…</si> entries handled below.
    const itemRegex = /<(?:(?:[A-Za-z_][\w.-]*):)?si\b[^>]*>([\s\S]*?)<\/(?:(?:[A-Za-z_][\w.-]*):)?si\s*>|<(?:(?:[A-Za-z_][\w.-]*):)?si\b[^>]*\/>/gi;
    let itemMatch: RegExpExecArray | null;
    while ((itemMatch = itemRegex.exec(xml)) !== null) {
        const full = itemMatch[0];
        if (full.endsWith('/>') && itemMatch[1] === undefined) {
            result.push('');
            continue;
        }
        let content = itemMatch[1] ?? '';
        // Skip phonetic runs entirely (their <t> is a reading hint, not cell text).
        content = content.replace(/<(?:(?:[A-Za-z_][\w.-]*):)?rPh\b[^>]*>[\s\S]*?<\/(?:(?:[A-Za-z_][\w.-]*):)?rPh\s*>/gi, '');
        // Handle self-closing rPh as well.
        content = content.replace(/<(?:(?:[A-Za-z_][\w.-]*):)?rPh\b[^>]*\/>/gi, '');
        const textRegex = /<(?:(?:[A-Za-z_][\w.-]*):)?t\b[^>]*>([\s\S]*?)<\/(?:(?:[A-Za-z_][\w.-]*):)?t\s*>/gi;
        let value = '';
        let textMatch: RegExpExecArray | null;
        while ((textMatch = textRegex.exec(content)) !== null) {
            value += textMatch[1];
        }
        result.push(unescapeXmlText(value));
    }
    return result;
}

/**
 * Buffer-based equivalent of {@link parseSharedStringsXml} for the common
 * Excel markup: unprefixed `<si>`/`<t>`, plain or rich runs, self-closing
 * `<si/>`, and `<rPh>` blocks that must be skipped. Falls back to the regex
 * parser for anything it does not recognise (CDATA, comments, prefixes,
 * attributes on `<si>`).
 */
export function parseSharedStringsBuffer(buf: Buffer): string[] {
    const first = buf.indexOf('<si');
    if (first === -1) return parseSharedStringsXml(buf.toString('utf8'));
    // The declaration is fine; comments/doctype/prefixed markup are not.
    const declarationEnd = buf.indexOf('?>');
    if (buf.indexOf('<!') !== -1 ||
        (declarationEnd !== -1 && buf.indexOf('<?', declarationEnd + 2) !== -1)) {
        return parseSharedStringsXml(buf.toString('utf8'));
    }

    const result: string[] = [];
    let pos = 0;
    while (true) {
        const open = buf.indexOf('<si', pos);
        if (open === -1) break;
        const next = buf[open + 3];
        if (next === 0x2f) { // <si/>
            result.push('');
            pos = open + 5;
            continue;
        }
        if (next !== 0x3e) {
            // Attributes or prefixed name: leave it to the regex parser.
            return parseSharedStringsXml(buf.toString('utf8'));
        }
        const contentStart = open + 4;
        const close = buf.indexOf('</si>', contentStart);
        if (close === -1) {
            return parseSharedStringsXml(buf.toString('utf8'));
        }
        let contentPos = contentStart;
        let value = '';
        let sawText = false;
        while (contentPos < close) {
            const lt = buf.indexOf('<', contentPos);
            if (lt === -1 || lt >= close) break;
            if (buf[lt + 1] === 0x72 && buf[lt + 2] === 0x50 && buf[lt + 3] === 0x68) {
                // <rPh ...>…</rPh> or <rPh .../>
                const tagEnd = buf.indexOf('>', lt);
                if (tagEnd === -1) return parseSharedStringsXml(buf.toString('utf8'));
                if (buf[tagEnd - 1] === 0x2f) {
                    contentPos = tagEnd + 1;
                } else {
                    const rphEnd = buf.indexOf('</rPh>', tagEnd + 1);
                    if (rphEnd === -1) return parseSharedStringsXml(buf.toString('utf8'));
                    contentPos = rphEnd + 6;
                }
                continue;
            }
            if (buf[lt + 1] === 0x74 && (
                buf[lt + 2] === 0x3e || buf[lt + 2] === 0x20 ||
                buf[lt + 2] === 0x09 || buf[lt + 2] === 0x0a || buf[lt + 2] === 0x0d
            )) {
                const tagEnd = buf.indexOf('>', lt);
                if (tagEnd === -1 || tagEnd >= close) return parseSharedStringsXml(buf.toString('utf8'));
                const textEnd = buf.indexOf('</t>', tagEnd + 1);
                if (textEnd === -1 || textEnd > close) return parseSharedStringsXml(buf.toString('utf8'));
                if (sawText) {
                    value += buf.toString('utf8', tagEnd + 1, textEnd);
                } else {
                    value = buf.toString('utf8', tagEnd + 1, textEnd);
                    sawText = true;
                }
                contentPos = textEnd + 4;
                continue;
            }
            // <r>, </r>, <rPr> and its children: skip the tag.
            const tagEnd = buf.indexOf('>', lt);
            if (tagEnd === -1) return parseSharedStringsXml(buf.toString('utf8'));
            contentPos = tagEnd + 1;
        }
        result.push(unescapeXmlText(value));
        pos = close + 5;
    }
    return result;
}
