/** Small measurement helpers shared by the benchmark runner. */

export interface MemorySample {
    peakRssMb: number;
    peakHeapMb: number;
}

export class MemoryMeter {
    private timer: NodeJS.Timeout | null = null;
    private peakRss = 0;
    private peakHeap = 0;

    start(): void {
        this.sample();
        this.timer = setInterval(() => this.sample(), 10);
        // Do not keep the process alive just for the sampler.
        this.timer.unref?.();
    }

    stop(): MemorySample {
        if (this.timer) clearInterval(this.timer);
        this.timer = null;
        this.sample();
        return { peakRssMb: this.peakRss, peakHeapMb: this.peakHeap };
    }

    private sample(): void {
        const usage = process.memoryUsage();
        this.peakRss = Math.max(this.peakRss, usage.rss / 1024 / 1024);
        this.peakHeap = Math.max(this.peakHeap, usage.heapUsed / 1024 / 1024);
    }
}

export function median(values: number[]): number {
    if (values.length === 0) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

export function trimOutliers(values: number[]): number[] {
    if (values.length <= 4) return values;
    const sorted = [...values].sort((a, b) => a - b);
    return sorted.slice(1, -1);
}

export function nowMs(): number {
    return Number(process.hrtime.bigint()) / 1e6;
}

export function gcIfAvailable(): void {
    const gc = (globalThis as { gc?: () => void }).gc;
    if (typeof gc === 'function') {
        gc();
        gc();
    }
}

export async function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
    let timer: NodeJS.Timeout | undefined;
    try {
        return await Promise.race([
            promise,
            new Promise<T>((_, reject) => {
                timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms} ms`)), ms);
                timer.unref?.();
            }),
        ]);
    } finally {
        if (timer) clearTimeout(timer);
    }
}
