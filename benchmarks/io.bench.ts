/**
 * End-to-end I/O benchmark with one writer/reader case per child process.
 *
 * This is the Etap 0 baseline tool: it reports encode/finalize time, peak RSS
 * and file size for the same deterministic datasets the project uses in
 * comparisons with xlspy. Run it with `npm run benchmark:io` (or
 * `npx ts-node benchmarks/io.bench.ts --quick`).
 */
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { BenchCase, defaultCases, makeHeaders, makeRows, expectedNonNullCount } from './datasets';
import { MemoryMeter, gcIfAvailable, median, nowMs, trimOutliers } from './metrics';

const RESULTS_DIR = path.resolve(__dirname, '..', 'benchmark_results');
const DEFAULT_RUNS = 3;

interface CaseResult {
    case: BenchCase;
    totalMs: number;
    writeMs?: number;
    finalizeMs?: number;
    openMs?: number;
    iterateMs?: number;
    sizeBytes: number;
    peakRssMb: number;
    peakHeapMb: number;
    rows: number;
    cells: number;
    cellsPerSecond: number;
    notes?: string;
}

interface AggregateResult extends CaseResult {
    medianMs: number;
    minMs: number;
    maxMs: number;
}

function tempFileFor(c: BenchCase): string {
    return path.join(os.tmpdir(), `justybase-bench-${c.id}.${c.format}`);
}

async function runWrite(c: BenchCase): Promise<CaseResult> {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { XlsbWriter } = require('../src/XlsbWriter') as typeof import('../src/XlsbWriter');
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const { XlsxWriter } = require('../src/XlsxWriter') as typeof import('../src/XlsxWriter');
    const data = makeRows(c.dataset, c.rows, c.cols);
    const headers = makeHeaders(c.cols);
    const file = tempFileFor(c);
    try {
        if (fs.existsSync(file)) fs.unlinkSync(file);
    } catch {
        // A stale temp file does not change the measurement.
    }

    gcIfAvailable();
    const meter = new MemoryMeter();
    meter.start();
    const writer = c.format === 'xlsx'
        ? new (XlsxWriter as any)(file, {
            compressionLevel: c.compressionLevel,
            useSharedStrings: c.sharedStrings,
            durable: false,
        })
        : new (XlsbWriter as any)(file, {
            compressionLevel: c.compressionLevel,
            useSharedStrings: c.sharedStrings,
            durable: false,
        });

    const writeStart = nowMs();
    if (c.mode === 'batch') {
        writer.addSheet('Bench');
        writer.writeSheet(data, headers, true);
    } else {
        writer.startSheet('Bench', c.cols, headers, {
            doAutofilter: true,
            sampleRows: data.slice(0, 100),
        });
        for (let row = 0; row < data.length; row++) writer.writeRow(data[row]);
        writer.endSheet();
    }
    const writeEnd = nowMs();
    await writer.finalize();
    const totalMs = nowMs() - writeStart;
    const memory = meter.stop();
    const sizeBytes = fs.existsSync(file) ? fs.statSync(file).size : 0;
    return {
        case: c,
        totalMs,
        writeMs: writeEnd - writeStart,
        finalizeMs: totalMs - (writeEnd - writeStart),
        sizeBytes,
        ...memory,
        rows: c.rows,
        cells: c.rows * c.cols,
        cellsPerSecond: (c.rows * c.cols) / Math.max(totalMs / 1000, 1e-9),
        notes: c.notes,
    };
}

async function ensureSourceFile(c: BenchCase): Promise<string> {
    const file = tempFileFor(c);
    if (fs.existsSync(file) && fs.statSync(file).size > 0) return file;
    const { XlsbWriter } = require('../src/XlsbWriter') as typeof import('../src/XlsbWriter');
    const { XlsxWriter } = require('../src/XlsxWriter') as typeof import('../src/XlsxWriter');
    const data = makeRows(c.dataset, c.rows, c.cols);
    const headers = makeHeaders(c.cols);
    const writer = c.format === 'xlsx'
        ? new (XlsxWriter as any)(file, { durable: false })
        : new (XlsbWriter as any)(file, { durable: false });
    writer.addSheet('Bench');
    writer.writeSheet(data, headers, true);
    await writer.finalize();
    return file;
}

async function runRead(c: BenchCase): Promise<CaseResult> {
    const { XlsbReader } = require('../src/XlsbReader') as typeof import('../src/XlsbReader');
    const { XlsxReader } = require('../src/XlsxReader') as typeof import('../src/XlsxReader');
    const file = await ensureSourceFile(c);
    gcIfAvailable();
    const meter = new MemoryMeter();
    meter.start();
    const reader: any = c.format === 'xlsx' ? new XlsxReader() : new XlsbReader();
    const openStart = nowMs();
    await reader.open(file);
    if (typeof reader.selectSheet === 'function' && reader.getSheetNames?.().length) {
        await reader.selectSheet(0);
    }
    const openEnd = nowMs();
    let rows = 0;
    let cells = 0;
    const readRow = typeof reader.readRow === 'function' ? reader.readRow.bind(reader) : null;
    const rowSource = readRow
        ? async () => {
            while (true) {
                const row = await readRow();
                if (row === null) return null;
                return row;
            }
        }
        : null;
    let nonNull = 0;
    if (rowSource) {
        // `readRow` is synchronous in the new API; keep the loop local.
        while (true) {
            const row = readRow!();
            if (row === null) break;
            rows++;
            cells += row.length;
            for (const value of row) if (value !== null && value !== undefined) nonNull++;
        }
    } else {
        while (await reader.read()) {
            rows++;
            for (let col = 0; col < reader.fieldCount; col++) {
                const value = reader.getValue(col);
                if (value !== null && value !== undefined) nonNull++;
            }
            cells += reader.fieldCount;
        }
    }
    // Guard against readers that silently return blanks: the source workbook
    // was written from the same deterministic dataset, so the number of
    // non-null cells is known exactly.
    const expectedNonNull = expectedNonNullCount(c.dataset, c.rows, c.cols) + c.cols; // + header row
    if (nonNull !== expectedNonNull) {
        throw new Error(`${c.id}: read ${nonNull} non-null cells, expected ${expectedNonNull}`);
    }
    const iterateEnd = nowMs();
    await reader.close();
    const totalMs = iterateEnd - openStart;
    const memory = meter.stop();
    return {
        case: c,
        totalMs,
        openMs: openEnd - openStart,
        iterateMs: iterateEnd - openEnd,
        sizeBytes: fs.existsSync(file) ? fs.statSync(file).size : 0,
        ...memory,
        rows,
        cells,
        cellsPerSecond: cells / Math.max(totalMs / 1000, 1e-9),
        notes: c.notes,
    };
}

function parseCaseArg(): BenchCase | null {
    const index = process.argv.indexOf('--case');
    if (index === -1) return null;
    const raw = process.argv[index + 1];
    if (!raw) throw new Error('--case needs a JSON argument');
    return JSON.parse(Buffer.from(raw, 'base64').toString('utf8')) as BenchCase;
}

async function childMain(c: BenchCase): Promise<void> {
    let result: CaseResult;
    if (c.kind === 'write') result = await runWrite(c);
    else result = await runRead(c);
    process.stdout.write(`RESULT:${JSON.stringify(result)}\n`);
}

function encodeCase(c: BenchCase): string {
    return Buffer.from(JSON.stringify(c), 'utf8').toString('base64');
}

function runChild(c: BenchCase): CaseResult {
    const result = spawnSync(
        process.execPath,
        ['-r', 'ts-node/register', '--expose-gc', __filename, '--case', encodeCase(c)],
        {
            cwd: path.resolve(__dirname, '..'),
            encoding: 'utf8',
            timeout: 30 * 60 * 1000,
            maxBuffer: 16 * 1024 * 1024,
            env: { ...process.env, TS_NODE_TRANSPILE_ONLY: '1' },
        }
    );
    if (result.status !== 0) {
        throw new Error(`case ${c.id} failed (${result.status}):\n${result.stdout}\n${result.stderr}`);
    }
    const line = (result.stdout || '').split(/\r?\n/).reverse().find(l => l.startsWith('RESULT:'));
    if (!line) throw new Error(`case ${c.id} produced no RESULT line:\n${result.stdout}\n${result.stderr}`);
    return JSON.parse(line.slice('RESULT:'.length)) as CaseResult;
}

function aggregate(results: CaseResult[]): AggregateResult {
    const times = trimOutliers(results.map(r => r.totalMs));
    const last = results[results.length - 1];
    return {
        ...last,
        medianMs: median(times),
        minMs: Math.min(...times),
        maxMs: Math.max(...times),
    };
}

function printSummary(results: AggregateResult[]): void {
    const header = 'case'.padEnd(38) + 'median'.padStart(10) + 'min'.padStart(9) + 'max'.padStart(9) +
        'cells/s'.padStart(13) + 'RSS MB'.padStart(10) + 'file MB'.padStart(10);
    console.log('\n' + header);
    console.log('-'.repeat(header.length));
    for (const r of results) {
        console.log(
            r.case.id.padEnd(38) +
            `${r.medianMs.toFixed(1)}`.padStart(10) +
            `${r.minMs.toFixed(1)}`.padStart(9) +
            `${r.maxMs.toFixed(1)}`.padStart(9) +
            `${Math.round(r.cellsPerSecond).toLocaleString('en-US')}`.padStart(13) +
            r.peakRssMb.toFixed(1).padStart(10) +
            (r.sizeBytes / 1024 / 1024).toFixed(2).padStart(10)
        );
    }
}

let RUNS_PER_REPORT = DEFAULT_RUNS;

function writeReports(results: AggregateResult[], startedAt: Date): void {
    fs.mkdirSync(RESULTS_DIR, { recursive: true });
    const stamp = startedAt.toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const jsonPath = path.join(RESULTS_DIR, `io-${stamp}.json`);
    const mdPath = path.join(RESULTS_DIR, `io-${stamp}.md`);
    fs.writeFileSync(jsonPath, JSON.stringify({
        date: startedAt.toISOString(),
        node: process.version,
        platform: `${os.platform()} ${os.release()}`,
        cpu: os.cpus()[0]?.model ?? 'unknown',
        results,
    }, null, 2));

    const lines = [
        '# I/O benchmark',
        '',
        `- Date: ${startedAt.toISOString()}`,
        `- Node: ${process.version}`,
        `- Platform: ${os.platform()} ${os.release()}`,
        `- CPU: ${os.cpus()[0]?.model ?? 'unknown'}`,
        '',
        '| Case | Median ms | Min ms | Max ms | Cells/s | Peak RSS MB | File MB |',
        '|---|---:|---:|---:|---:|---:|---:|',
        ...results.map(r => `| ${r.case.id} | ${r.medianMs.toFixed(1)} | ${r.minMs.toFixed(1)} | ${r.maxMs.toFixed(1)} | ${Math.round(r.cellsPerSecond).toLocaleString('en-US')} | ${r.peakRssMb.toFixed(1)} | ${(r.sizeBytes / 1024 / 1024).toFixed(2)} |`),
        '',
        `Each case runs in its own process; the median is over ${RUNS_PER_REPORT} runs after trimming the slowest and fastest.`,
    ];
    fs.writeFileSync(mdPath, lines.join('\n') + '\n');
    console.log(`\nResults: ${jsonPath}\n         ${mdPath}`);
}

async function parentMain(): Promise<void> {
    const quick = process.argv.includes('--quick');
    const onlyIndex = process.argv.indexOf('--only');
    const only = onlyIndex === -1 ? null : process.argv[onlyIndex + 1];
    const baselineIndex = process.argv.indexOf('--baseline');
    const compareIndex = process.argv.indexOf('--check');
    const maxRegressionIndex = process.argv.indexOf('--max-regression');
    const runsIndex = process.argv.indexOf('--runs');
    const runCount = runsIndex === -1 ? DEFAULT_RUNS : Math.max(1, Number(process.argv[runsIndex + 1]));
    const maxRegressionPct = maxRegressionIndex === -1 ? 5 : Number(process.argv[maxRegressionIndex + 1]);
    const saveBaseline = process.argv.includes('--save-baseline');
    const baselinePath = baselineIndex !== -1
        ? path.resolve(process.argv[baselineIndex + 1])
        : path.join(RESULTS_DIR, 'baseline.json');
    const shouldCompare = compareIndex !== -1 || baselineIndex !== -1;
    let cases = defaultCases(quick);
    if (only) cases = cases.filter(c => c.id.includes(only));
    if (cases.length === 0) {
        console.error('No benchmark cases matched.');
        process.exitCode = 1;
        return;
    }
    RUNS_PER_REPORT = runCount;
    console.log(`Running ${cases.length} cases x ${runCount} runs (${quick ? 'quick' : 'full'})...`);
    const startedAt = new Date();
    const results: AggregateResult[] = [];
    for (const c of cases) {
        const runs: CaseResult[] = [];
        for (let run = 0; run < runCount; run++) {
            process.stdout.write(`  ${c.id} [${run + 1}/${runCount}] ... `);
            const result = runChild(c);
            runs.push(result);
            process.stdout.write(`${result.totalMs.toFixed(1)} ms\n`);
        }
        results.push(aggregate(runs));
    }
    printSummary(results);
    writeReports(results, startedAt);

    if (saveBaseline) {
        fs.mkdirSync(RESULTS_DIR, { recursive: true });
        fs.writeFileSync(path.join(RESULTS_DIR, 'baseline.json'), JSON.stringify({
            date: startedAt.toISOString(),
            node: process.version,
            platform: `${os.platform()} ${os.release()}`,
            cpu: os.cpus()[0]?.model ?? 'unknown',
            results,
        }, null, 2));
        console.log(`Baseline saved to ${path.join(RESULTS_DIR, 'baseline.json')}`);
    }

    if (shouldCompare) {
        if (!fs.existsSync(baselinePath)) {
            console.error(`Baseline not found: ${baselinePath}`);
            process.exitCode = 1;
            return;
        }
        const baseline = JSON.parse(fs.readFileSync(baselinePath, 'utf8')) as { results: AggregateResult[] };
        const byId = new Map(baseline.results.map(r => [r.case.id, r]));
        console.log(`\nRegression check vs ${baselinePath} (limit ${maxRegressionPct}%):`);
        const regressions: string[] = [];
        for (const result of results) {
            const previous = byId.get(result.case.id);
            if (!previous) continue;
            const change = ((result.medianMs - previous.medianMs) / previous.medianMs) * 100;
            const mark = change > maxRegressionPct ? 'REGRESSION' : 'ok';
            console.log(`  ${result.case.id.padEnd(38)} ${change >= 0 ? '+' : ''}${change.toFixed(1)}%  ${mark}`);
            if (change > maxRegressionPct) regressions.push(result.case.id);
        }
        if (regressions.length > 0) {
            console.error(`\n${regressions.length} case(s) exceeded the ${maxRegressionPct}% regression gate: ${regressions.join(', ')}`);
            process.exitCode = 1;
        }
    }
}

const childCase = parseCaseArg();
if (childCase) {
    childMain(childCase).catch(error => {
        console.error(error);
        process.exit(1);
    });
} else {
    parentMain().catch(error => {
        console.error(error);
        process.exit(1);
    });
}
