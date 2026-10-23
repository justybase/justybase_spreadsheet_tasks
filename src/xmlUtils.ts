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
    if (!/&|<|>|"|'/.test(str) && !hasInvalidControlChars(str)) {
        return str;
    }
    let out = str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
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
        result.push(unescapeXml(value));
    }
    return result;
}
