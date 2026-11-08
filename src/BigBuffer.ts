import { Buffer } from 'buffer';

export class BigBuffer {
    private chunkSize: number;
    private chunks: Buffer[];
    private currentBuffer: Buffer;
    private cursor: number;

    constructor(chunkSize: number = 65536) {
        this.chunkSize = chunkSize;
        this.chunks = [];
        this.currentBuffer = Buffer.alloc(chunkSize);
        this.cursor = 0;
    }

    ensureCapacity(size: number): void {
        if (this.cursor + size > this.chunkSize) {
            this._flush();
        }
    }

    private _ensureCapacity(size: number): void {
        this.ensureCapacity(size);
    }

    private _flush(): void {
        if (this.cursor > 0) {
            this.chunks.push(this.currentBuffer.subarray(0, this.cursor));
            this.currentBuffer = Buffer.alloc(this.chunkSize);
            this.cursor = 0;
        }
    }

    write(buffer: Uint8Array): void {
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

    writeByte(val: number): void {
        this._ensureCapacity(1);
        this.currentBuffer[this.cursor] = val;
        this.cursor++;
    }

    writeUnsafeByte(val: number): void {
        this.currentBuffer[this.cursor++] = val;
    }

    writeInt32LE(val: number): void {
        this._ensureCapacity(4);
        this.currentBuffer.writeInt32LE(val, this.cursor);
        this.cursor += 4;
    }

    writeUnsafeInt32LE(val: number): void {
        this.currentBuffer.writeInt32LE(val, this.cursor);
        this.cursor += 4;
    }

    writeDoubleLE(val: number): void {
        this._ensureCapacity(8);
        this.currentBuffer.writeDoubleLE(val, this.cursor);
        this.cursor += 8;
    }

    writeUnsafeDoubleLE(val: number): void {
        this.currentBuffer.writeDoubleLE(val, this.cursor);
        this.cursor += 8;
    }

    writeString(str: string): void {
        // Fast path for ASCII markup: write code units directly, no UTF-8 encode.
        if (this.cursor + str.length <= this.chunkSize) {
            let ascii = true;
            for (let i = 0; i < str.length; i++) {
                if (str.charCodeAt(i) >= 0x80) {
                    ascii = false;
                    break;
                }
            }
            if (ascii) {
                for (let i = 0; i < str.length; i++) {
                    this.currentBuffer[this.cursor++] = str.charCodeAt(i);
                }
                return;
            }
        }

        const byteLength = Buffer.byteLength(str, 'utf8');

        if (this.cursor + byteLength <= this.chunkSize) {
            this.cursor += this.currentBuffer.write(str, this.cursor, 'utf8');
            return;
        }

        this._flush();

        if (byteLength > this.chunkSize) {
            this.chunks.push(Buffer.from(str, 'utf8'));
        } else {
            this.cursor += this.currentBuffer.write(str, this.cursor, 'utf8');
        }
    }

    writeUtf16LE(str: string): void {
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
            this.chunks.push(Buffer.from(str, 'utf16le'));
        } else {
            for (let i = 0; i < str.length; i++) {
                const code = str.charCodeAt(i);
                this.currentBuffer[this.cursor++] = code & 0xFF;
                this.currentBuffer[this.cursor++] = (code >> 8) & 0xFF;
            }
        }
    }

    getChunks(): Buffer[] {
        if (this.cursor > 0) {
            this.chunks.push(this.currentBuffer.subarray(0, this.cursor));
            this.currentBuffer = Buffer.alloc(this.chunkSize);
            this.cursor = 0;
        }
        return this.chunks;
    }

    /** Overwrite a previously written 32-bit integer at an absolute offset. */
    patchInt32LE(offset: number, value: number): boolean {
        let position = 0;
        for (const chunk of this.chunks) {
            if (offset >= position && offset + 4 <= position + chunk.length) {
                chunk.writeInt32LE(value, offset - position);
                return true;
            }
            position += chunk.length;
        }
        if (offset >= position && offset + 4 <= position + this.cursor) {
            this.currentBuffer.writeInt32LE(value, offset - position);
            return true;
        }
        return false;
    }

    reset(): void {
        this.chunks = [];
        this.cursor = 0;
    }
}

export default BigBuffer;
