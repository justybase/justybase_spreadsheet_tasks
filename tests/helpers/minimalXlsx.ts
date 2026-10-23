import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';

export function mkTempDir(prefix = 'spreadsheet-helper-'): string {
    return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

export function addText(zip: AdmZip, name: string, text: string): void {
    zip.addFile(name, Buffer.from(text, 'utf8'));
}

export interface MinimalXlsxOptions {
    workbookXml?: string;
    relsXml?: string;
    sheets?: Array<{ name: string; xml: string; relTarget?: string; rId?: string }>;
    sharedStringsXml?: string | null;
    stylesXml?: string | null;
    extraFiles?: Record<string, string | Buffer>;
}

const DEFAULT_WORKBOOK = (sheets: Array<{ name: string; rId: string }>): string =>
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="${s.rId}"/>`).join('')}</sheets></workbook>`;

const DEFAULT_RELS = (targets: Array<{ rId: string; target: string; type?: string }>): string =>
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${targets.map(t => `<Relationship Id="${t.rId}" Type="${t.type ?? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'}" Target="${t.target}"/>`).join('')}</Relationships>`;

export function createMinimalXlsx(filePath: string, opts: MinimalXlsxOptions = {}): void {
    const zip = new AdmZip();
    addText(zip, '[Content_Types].xml', '<?xml version="1.0"?><Types/>');

    const sheets = opts.sheets ?? [{
        name: 'data1',
        xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        rId: 'rId1',
    }];

    const wbXml = opts.workbookXml ?? DEFAULT_WORKBOOK(sheets.map(s => ({ name: s.name, rId: s.rId ?? 'rId1' })));
    addText(zip, 'xl/workbook.xml', wbXml);

    const relsTargets = sheets.map((s, i) => ({
        rId: s.rId ?? `rId${i + 1}`,
        target: s.relTarget ?? `worksheets/sheet${i + 1}.xml`,
    }));
    const relsXml = opts.relsXml ?? DEFAULT_RELS(relsTargets);
    addText(zip, 'xl/_rels/workbook.xml.rels', relsXml);

    sheets.forEach((s, i) => {
        addText(zip, `xl/worksheets/sheet${i + 1}.xml`, s.xml);
    });

    if (opts.sharedStringsXml === undefined) {
        // no SST by default
    } else if (opts.sharedStringsXml !== null) {
        addText(zip, 'xl/sharedStrings.xml', opts.sharedStringsXml);
    }

    if (opts.stylesXml) {
        addText(zip, 'xl/styles.xml', opts.stylesXml);
    }

    if (opts.extraFiles) {
        for (const [name, content] of Object.entries(opts.extraFiles)) {
            if (typeof content === 'string') addText(zip, name, content);
            else zip.addFile(name, content);
        }
    }

    zip.writeZip(filePath);
}

export function zipText(filePath: string, name: string): string {
    const zip = new AdmZip(filePath);
    const entry = zip.getEntry(name);
    if (!entry) throw new Error(`missing ZIP member ${name}`);
    return entry.getData().toString('utf8');
}

export function zipHas(filePath: string, name: string): boolean {
    const zip = new AdmZip(filePath);
    return !!zip.getEntry(name);
}
