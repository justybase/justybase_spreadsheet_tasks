/**
 * Minimal byte sink implemented by both the Node `BigBuffer` and the browser
 * `BrowserBigBuffer`. The shared writer cores only depend on this interface, so
 * the same serialization code runs in Node and in the browser.
 */
export interface ByteSink {
    ensureCapacity(size: number): void;
    writeByte(val: number): void;
    writeUnsafeByte(val: number): void;
    writeInt32LE(val: number): void;
    writeUnsafeInt32LE(val: number): void;
    writeDoubleLE(val: number): void;
    writeUnsafeDoubleLE(val: number): void;
    write(bytes: Uint8Array): void;
    writeString(str: string): void;
    writeUtf16LE(str: string): void;
}

/** Sink used by the XLSX core; XML is always written as UTF-8 text. */
export interface TextByteSink {
    writeString(str: string): void;
}
