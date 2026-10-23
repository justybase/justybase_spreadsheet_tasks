export class BrowserBigBuffer {
    chunkSize: number;
    chunks: Uint8Array<ArrayBuffer>[];
    currentBuffer: Uint8Array<ArrayBuffer>;
    dataView: DataView;
    cursor: number;
    private _textEncoder: InstanceType<typeof TextEncoder>;
    /**
     * @param {number} [chunkSize=65536]
     */
    constructor(chunkSize: number = 65536) {
        this.chunkSize = chunkSize;
        /** @type {Uint8Array[]} */
        this.chunks = [];
        this.currentBuffer = new Uint8Array(chunkSize);
        this.dataView = new DataView(this.currentBuffer.buffer);
        this.cursor = 0;
        this._textEncoder = new TextEncoder();
    }

    /**
     * @param {number} size
     */
    ensureCapacity(size: number) {
        if (this.cursor + size > this.chunkSize) {
            this._flush();
        }
    }

    _flush() {
        if (this.cursor > 0) {
            this.chunks.push(this.currentBuffer.slice(0, this.cursor));
            this.currentBuffer = new Uint8Array(this.chunkSize);
            this.dataView = new DataView(this.currentBuffer.buffer);
            this.cursor = 0;
        }
    }

    /**
     * Write a Uint8Array to the buffer.
     * @param {Uint8Array} buffer
     */
    write(buffer: Uint8Array) {
        let len = buffer.length;
        let offset = 0;

        while (len > 0) {
            let available = this.chunkSize - this.cursor;
            if (available === 0) {
                this._flush();
                available = this.chunkSize;
            }

            const toWrite = Math.min(len, available);
            this.currentBuffer.set(buffer.subarray(offset, offset + toWrite), this.cursor);

            this.cursor += toWrite;
            offset += toWrite;
            len -= toWrite;
        }
    }

    /**
     * @param {number} val
     */
    writeByte(val: number) {
        this.ensureCapacity(1);
        this.currentBuffer[this.cursor] = val;
        this.cursor++;
    }

    /**
     * @param {number} val
     */
    writeUnsafeByte(val: number) {
        this.currentBuffer[this.cursor++] = val;
    }

    /**
     * @param {number} val
     */
    writeInt32LE(val: number) {
        this.ensureCapacity(4);
        this.dataView.setInt32(this.cursor, val, true);
        this.cursor += 4;
    }

    /**
     * @param {number} val
     */
    writeUnsafeInt32LE(val: number) {
        this.dataView.setInt32(this.cursor, val, true);
        this.cursor += 4;
    }

    /**
     * @param {number} val
     */
    writeDoubleLE(val: number) {
        this.ensureCapacity(8);
        this.dataView.setFloat64(this.cursor, val, true);
        this.cursor += 8;
    }

    /**
     * @param {number} val
     */
    writeUnsafeDoubleLE(val: number) {
        this.dataView.setFloat64(this.cursor, val, true);
        this.cursor += 8;
    }

    /**
     * Write a UTF-8 string.
     * @param {string} str
     */
    writeString(str: string) {
        const encoded = this._textEncoder.encode(str);
        const byteLength = encoded.length;

        if (this.cursor + byteLength <= this.chunkSize) {
            this.currentBuffer.set(encoded, this.cursor);
            this.cursor += byteLength;
            return;
        }

        this._flush();

        if (byteLength > this.chunkSize) {
            this.chunks.push(encoded);
        } else {
            this.currentBuffer.set(encoded, this.cursor);
            this.cursor += byteLength;
        }
    }

    /**
     * Write a string as UTF-16LE.
     * @param {string} str
     */
    writeUtf16LE(str: string) {
        const byteLength = str.length * 2;

        if (this.cursor + byteLength <= this.chunkSize) {
            for (let i = 0; i < str.length; i++) {
                const code = str.charCodeAt(i);
                this.currentBuffer[this.cursor++] = code & 0xFF;
                this.currentBuffer[this.cursor++] = (code >> 8) & 0xFF;
            }
            return;
        }

        this._flush();
        if (byteLength > this.chunkSize) {
            const buf = new Uint8Array(byteLength);
            for (let i = 0; i < str.length; i++) {
                const code = str.charCodeAt(i);
                buf[i * 2] = code & 0xFF;
                buf[i * 2 + 1] = (code >> 8) & 0xFF;
            }
            this.chunks.push(buf);
        } else {
            for (let i = 0; i < str.length; i++) {
                const code = str.charCodeAt(i);
                this.currentBuffer[this.cursor++] = code & 0xFF;
                this.currentBuffer[this.cursor++] = (code >> 8) & 0xFF;
            }
        }
    }

    /**
     * Flush and return all chunks.
     * @returns {Uint8Array[]}
     */
    getChunks(): Uint8Array<ArrayBuffer>[] {
        if (this.cursor > 0) {
            this.chunks.push(this.currentBuffer.slice(0, this.cursor));
            this.currentBuffer = new Uint8Array(this.chunkSize);
            this.dataView = new DataView(this.currentBuffer.buffer);
            this.cursor = 0;
        }
        return this.chunks;
    }

    /**
     * Concatenate all chunks into a single Uint8Array.
     * @returns {Uint8Array}
     */
    toUint8Array(): Uint8Array<ArrayBuffer> {
        const chunks = this.getChunks();
        let totalLen = 0;
        for (const c of chunks) totalLen += c.length;

        const result = new Uint8Array(totalLen);
        let offset = 0;
        for (const c of chunks) {
            result.set(c, offset);
            offset += c.length;
        }
        return result;
    }

    reset() {
        this.chunks = [];
        this.cursor = 0;
    }
}

export default BrowserBigBuffer;
