/**
 * Runs the module script of browser/demo.html or browser/test.html outside a
 * browser and writes every generated download to a directory.
 *
 * Usage: node run-browser-page.mjs <demo|test> <html-file> <out-dir> <bundle>
 *
 * This is a plain .mjs child process on purpose: ts-node compiles the test
 * runner to CommonJS, which cannot import the page's ES module.
 */
import fs from 'fs';
import path from 'path';
import { pathToFileURL } from 'url';

const [mode, htmlFile, outDir, bundleFile] = process.argv.slice(2);
if (!mode || !htmlFile || !outDir || !bundleFile) {
    console.error('usage: node run-browser-page.mjs <demo|test> <html> <out-dir> <bundle>');
    process.exit(2);
}

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

const bundleUrl = pathToFileURL(path.resolve(bundleFile)).href;
const html = fs.readFileSync(htmlFile, 'utf8');
const match = /<script type="module">([\s\S]*?)<\/script>/.exec(html);
if (!match) {
    console.error(`no module script in ${htmlFile}`);
    process.exit(2);
}
const modulePath = path.join(outDir, `page-${process.pid}.mjs`);
fs.writeFileSync(modulePath, match[1].replace(
    /from\s+['"]\.\/browser-spreadsheet\.js['"]/g,
    `from '${bundleUrl}'`,
));

const elements = new Map();
const downloads = new Map();
let pendingBlob = null;

function makeElement(name) {
    return {
        name,
        style: {},
        onclick: null,
        innerText: '',
        innerHTML: '',
        appendChild(child) {
            const text = child?.innerText || child?.innerHTML || '';
            if (text) console.log(text);
        },
        removeChild() {},
        click() {
            if (this.download && pendingBlob) downloads.set(this.download, pendingBlob);
        },
    };
}

globalThis.document = {
    getElementById(id) {
        if (!elements.has(id)) elements.set(id, makeElement(id));
        return elements.get(id);
    },
    createElement: makeElement,
    body: { appendChild() {}, removeChild() {} },
};
globalThis.URL.createObjectURL = blob => {
    pendingBlob = blob;
    return 'blob:mock';
};
globalThis.URL.revokeObjectURL = () => {};

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

await import(pathToFileURL(modulePath).href);

if (mode === 'demo') {
    const buttons = [
        'btnSimpleXlsx', 'btnSimpleXlsb',
        'btnLargeXlsx', 'btnLargeXlsb',
        'btnMultiXlsx', 'btnMultiXlsb',
        'btnFmtXlsx', 'btnFmtXlsb',
        'btnAllXlsx', 'btnAllXlsb',
        'btnShowcaseXlsx', 'btnShowcaseXlsb', 'btnUltimateXlsx', 'btnUltimateXlsb',
    ];
    for (const id of buttons) {
        const element = elements.get(id);
        if (!element || typeof element.onclick !== 'function') {
            console.error(`button ${id} has no handler`);
            process.exit(2);
        }
        element.onclick();
        await sleep(50);
    }
} else if (mode === 'test') {
    // The test page starts async work on load and enables the download button
    // when every writer has produced its blob.
    const deadline = Date.now() + 120_000;
    let button = elements.get('btnDownload');
    while (Date.now() < deadline && (!button || button.style.display !== 'block')) {
        await sleep(100);
        button = elements.get('btnDownload');
    }
    if (!button || typeof button.onclick !== 'function') {
        console.error('test page did not enable the download button');
        process.exit(2);
    }
    const before = downloads.size;
    button.onclick();
    let stable = 0;
    let previous = -1;
    while (Date.now() < deadline && stable < 5) {
        await sleep(100);
        if (downloads.size === previous && downloads.size > before) {
            stable++;
        } else {
            stable = 0;
        }
        previous = downloads.size;
    }
    if (downloads.size <= before) {
        console.error('test page produced no downloads');
        process.exit(2);
    }
} else {
    console.error(`unknown mode ${mode}`);
    process.exit(2);
}

for (const [name, blob] of downloads) {
    const file = path.join(outDir, name);
    fs.writeFileSync(file, Buffer.from(await blob.arrayBuffer()));
    console.log(`WROTE ${name} ${fs.statSync(file).size}`);
}
process.exit(0);
