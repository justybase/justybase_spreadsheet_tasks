export class BrowserZip {
    files: { name: string; data: Uint8Array<ArrayBuffer> }[] = [];
    private static _crc32Table: Uint32Array | null = null;

    constructor() {
        this.files = [];
    }

    /**
     * Add a file entry (string or Uint8Array).
     * @param {string} name - Path inside the ZIP (e.g. "xl/workbook.xml")
     * @param {string | Uint8Array | Uint8Array[]} data - File contents
     */
    addFile(name: string, data: string | Uint8Array<ArrayBuffer> | Uint8Array<ArrayBuffer>[]): void {
        let bytes: Uint8Array<ArrayBuffer>;
        if (typeof data === 'string') {
            bytes = new TextEncoder().encode(data);
        } else if (Array.isArray(data)) {
            // Array of Uint8Array chunks — concatenate
            let totalLen = 0;
            for (const c of data) totalLen += c.length;
            bytes = new Uint8Array(totalLen);
            let offset = 0;
            for (const c of data) {
                bytes.set(c, offset);
                offset += c.length;
            }
        } else {
            bytes = data;
        }
        this.files.push({ name, data: bytes });
    }

    /**
     * Generate the ZIP file as a Blob.
     * @returns {Blob}
     */
    toBlob(): Blob {
        const parts: (Uint8Array<ArrayBuffer> | Blob)[] = [];
        const centralDir: Uint8Array<ArrayBuffer>[] = [];
        let localOffset = 0;

        for (const file of this.files) {
            const nameBytes = new TextEncoder().encode(file.name);
            const crc = this._crc32(file.data);

            // Local file header (30 + nameLen bytes)
            const localHeader = new Uint8Array(30 + nameBytes.length);
            const lhView = new DataView(localHeader.buffer);

            lhView.setUint32(0, 0x04034b50, true);   // Local file header signature
            lhView.setUint16(4, 20, true);             // Version needed to extract (2.0)
            lhView.setUint16(6, 0, true);              // General purpose bit flag
            lhView.setUint16(8, 0, true);              // Compression method: STORE
            lhView.setUint16(10, 0, true);             // Last mod time
            lhView.setUint16(12, 0, true);             // Last mod date
            lhView.setUint32(14, crc, true);           // CRC-32
            lhView.setUint32(18, file.data.length, true); // Compressed size
            lhView.setUint32(22, file.data.length, true); // Uncompressed size
            lhView.setUint16(26, nameBytes.length, true);  // File name length
            lhView.setUint16(28, 0, true);             // Extra field length
            localHeader.set(nameBytes, 30);

            parts.push(localHeader);
            parts.push(file.data);

            // Central directory entry (46 + nameLen bytes)
            const cdEntry = new Uint8Array(46 + nameBytes.length);
            const cdView = new DataView(cdEntry.buffer);

            cdView.setUint32(0, 0x02014b50, true);    // Central directory header signature
            cdView.setUint16(4, 20, true);             // Version made by
            cdView.setUint16(6, 20, true);             // Version needed to extract
            cdView.setUint16(8, 0, true);              // General purpose bit flag
            cdView.setUint16(10, 0, true);             // Compression method: STORE
            cdView.setUint16(12, 0, true);             // Last mod time
            cdView.setUint16(14, 0, true);             // Last mod date
            cdView.setUint32(16, crc, true);           // CRC-32
            cdView.setUint32(20, file.data.length, true); // Compressed size
            cdView.setUint32(24, file.data.length, true); // Uncompressed size
            cdView.setUint16(28, nameBytes.length, true);  // File name length
            cdView.setUint16(30, 0, true);             // Extra field length
            cdView.setUint16(32, 0, true);             // File comment length
            cdView.setUint16(34, 0, true);             // Disk number start
            cdView.setUint16(36, 0, true);             // Internal file attributes
            cdView.setUint32(38, 0, true);             // External file attributes
            cdView.setUint32(42, localOffset, true);   // Relative offset of local header
            cdEntry.set(nameBytes, 46);

            centralDir.push(cdEntry);

            localOffset += localHeader.length + file.data.length;
        }

        // Write central directory
        const cdStartOffset = localOffset;
        let cdSize = 0;
        for (const entry of centralDir) {
            parts.push(entry);
            cdSize += entry.length;
        }

        // End of central directory record (22 bytes)
        const eocd = new Uint8Array(22);
        const eocdView = new DataView(eocd.buffer);
        eocdView.setUint32(0, 0x06054b50, true);       // End of central dir signature
        eocdView.setUint16(4, 0, true);                 // Number of this disk
        eocdView.setUint16(6, 0, true);                 // Disk where central dir starts
        eocdView.setUint16(8, this.files.length, true); // Number of entries on this disk
        eocdView.setUint16(10, this.files.length, true);// Total number of entries
        eocdView.setUint32(12, cdSize, true);           // Size of central directory
        eocdView.setUint32(16, cdStartOffset, true);    // Offset of start of central directory
        eocdView.setUint16(20, 0, true);                // Comment length

        parts.push(eocd);

        return new Blob(parts, { type: 'application/octet-stream' });
    }

    /**
     * CRC-32 calculation.
     * @param {Uint8Array} data
     * @returns {number}
     */
    _crc32(data: Uint8Array): number {
        if (!BrowserZip._crc32Table) {
            BrowserZip._crc32Table = new Uint32Array(256);
            for (let i = 0; i < 256; i++) {
                let c = i;
                for (let j = 0; j < 8; j++) {
                    c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
                }
                BrowserZip._crc32Table[i] = c;
            }
        }
        const table = BrowserZip._crc32Table;

        let crc = 0xFFFFFFFF;
        for (let i = 0; i < data.length; i++) {
            crc = table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
        }
        return (crc ^ 0xFFFFFFFF) >>> 0;
    }
}

/** @type {Uint32Array | null} */

export default BrowserZip;
