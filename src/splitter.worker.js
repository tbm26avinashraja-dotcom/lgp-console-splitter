import {
  ZipReader,
  ZipWriter,
  BlobReader,
  TextReader,
  Writer
} from '@zip.js/zip.js';

class WritableStreamWriter extends Writer {
  constructor(stream) {
    super();
    this.writer = stream.getWriter();
  }

  async writeUint8Array(array) {
    await this.writer.write(array);
  }

  async close() {
    await this.writer.close();
  }
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
const XML_END = '</sheetData></worksheet>';
const STATIC = {
  '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="LGP Console" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
};
const encoder = new TextEncoder();
const notify = (label, detail, percent) => self.postMessage({ type: 'progress', label, detail, percent });
const decodeXml = s => String(s ?? '').replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const escapeXml = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
const norm = s => String(s ?? '').normalize('NFKC').trim().replace(/\s+/g, ' ').toUpperCase();
const safeName = s => String(s).replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 85) || 'UNASSIGNED';
function columnIndex(letters) { let n = 0; for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
function cellParts(row) {
  const out = [];
  const re = /<c\b([^>]*?)(?:\s*\/>|>([\s\S]*?)<\/c\s*>)/g;
  let m;
  while ((m = re.exec(row))) {
    const ref = /\br="([A-Z]+)(\d+)"/.exec(m[1]);
    if (ref) out.push({ index: columnIndex(ref[1]), col: ref[1], attrs: m[1], body: m[2] || '' });
  }
  return out;
}
function extractCell(cell, shared) {
  const type = /\bt="([^"]+)"/.exec(cell.attrs)?.[1] || '';
  const raw = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(cell.body)?.[1];
  if (type === 's') {
    const idx = Number(raw);
    if (!Number.isInteger(idx) || idx < 0 || idx >= shared.length) throw new Error(`Invalid shared string index ${raw}.`);
    return shared[idx];
  }
  if (type === 'inlineStr') {
    return [...cell.body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m => decodeXml(m[1])).join('');
  }
  return decodeXml(raw ?? '');
}
function cleanRow(row, number, shared) {
  const cells = cellParts(row);
  let xml = `<row r="${number}">`;
  for (const cell of cells) {
    const type = /\bt="([^"]+)"/.exec(cell.attrs)?.[1] || '';
    const value = extractCell(cell, shared);
    const ref = `${cell.col}${number}`;
    if (type === 's' || type === 'inlineStr' || type === 'str' || type === 'e') {
      xml += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
    } else if (type === 'b') {
      xml += `<c r="${ref}" t="b"><v>${value === '1' ? '1' : '0'}</v></c>`;
    } else if (value !== '') {
      // Preserve stored numeric values. Do not carry formulas or style indices into the new workbook.
      if (!Number.isFinite(Number(value))) throw new Error(`Unexpected non-numeric cell ${ref}.`);
      xml += `<c r="${ref}" t="n"><v>${value}</v></c>`;
    }
  }
  return xml + '</row>';
}
function parseSharedStrings(xml) {
  const values = [];
  for (const si of xml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)) {
    values.push([...si[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m => decodeXml(m[1])).join(''));
  }
  return values;
}
async function run(file) {
  if (!file) throw new Error('Select an XLSX workbook.');
  if (!self.navigator?.storage?.getDirectory) throw new Error('Browser storage is unavailable. Use a recent Chrome or Edge browser on HTTPS.');
  const root = await self.navigator.storage.getDirectory();
  const dirName = `lgp-split-${Date.now()}`;
  const dir = await root.getDirectoryHandle(dirName, { create: true });
  const reader = new ZipReader(new BlobReader(file));
  const streams = new Set();
  try {
    const entries = await reader.getEntries();
    const sheet = entries.find(e => e.filename === 'xl/worksheets/sheet1.xml') || entries.find(e => /^xl\/worksheets\/sheet\d+\.xml$/.test(e.filename));
    if (!sheet) throw new Error('No worksheet found in this XLSX file.');
    const sharedEntry = entries.find(e => e.filename === 'xl/sharedStrings.xml');
    let shared = [];
    if (sharedEntry) {
      notify('Reading shared strings', 'Preparing Excel text values', 2);
      const text = await sharedEntry.getData(new (await import('@zip.js/zip.js')).TextWriter());
      shared = parseSharedStrings(text);
    }
    const groups = new Map();
    const usedNames = new Set();
    let headers = null, partyIndex = -1, headerXml = '', rows = 0, blankParties = 0;
    async function groupFor(party) {
      const key = norm(party) || '__UNASSIGNED__';
      if (groups.has(key)) return groups.get(key);
      let name = safeName(key === '__UNASSIGNED__' ? 'UNASSIGNED' : key), base = name, i = 2;
      while (usedNames.has(name.toLowerCase())) name = `${base.slice(0, 75)}_${i++}`;
      usedNames.add(name.toLowerCase());
      const handle = await dir.getFileHandle(`party_${groups.size}.xml`, { create: true });
      const stream = await handle.createWritable(); streams.add(stream);
      await stream.write(XML_HEAD + headerXml);
      const g = { key, name, handle, stream, rows: 0, buffer: '' };
      groups.set(key, g);
      return g;
    }
    async function accept(row) {
      if (!headers) {
        const cells = cellParts(row);
        headers = [];
        for (const cell of cells) headers[cell.index] = extractCell(cell, shared);
        partyIndex = headers.findIndex(h => norm(h).replace(/\s+/g, '_') === 'PARTY_NAME');
        if (partyIndex < 0) throw new Error('Required column "Party Name" (or party_name) was not found.');
        headerXml = cleanRow(row, 1, shared);
        return;
      }
      const cell = cellParts(row).find(c => c.index === partyIndex);
      const party = cell ? extractCell(cell, shared) : '';
      if (!String(party).trim()) blankParties++;
      const g = await groupFor(party);
      g.rows++; rows++;
      g.buffer += cleanRow(row, g.rows + 1, shared);
      if (g.buffer.length >= 256 * 1024) { await g.stream.write(g.buffer); g.buffer = ''; }
      if (rows % 10000 === 0) notify('Grouping transactions', `${rows.toLocaleString()} rows, ${groups.size} parties`, Math.min(65, 5 + Math.round(rows / 450000 * 60)));
    }
    notify('Reading workbook', 'Streaming worksheet data', 4);
    let pending = '', inData = false;
    const decoder = new TextDecoder();
    async function parseChunk(chunk) {
      pending += decoder.decode(chunk, { stream: true });
      if (!inData) {
        const m = /<sheetData(?:\s[^>]*)?>/.exec(pending);
        if (!m) { if (pending.length > 2_000_000) throw new Error('Could not locate worksheet data.'); return; }
        pending = pending.slice(m.index + m[0].length);
        inData = true;
      }
      while (true) {
        const start = /<row(?:\s|>)/.exec(pending);
        if (!start) { pending = pending.slice(-128); return; }
        if (start.index) pending = pending.slice(start.index);
        const end = pending.indexOf('</row>');
        if (end < 0) {
          if (pending.length > 20_000_000) throw new Error('A worksheet row is unexpectedly large.');
          return;
        }
        const row = pending.slice(0, end + 6);
        pending = pending.slice(end + 6);
        await accept(row);
      }
    }
    await sheet.getData(new WritableStreamWriter(new WritableStream({ write: parseChunk })));
    if (!headers || !rows) throw new Error('No transaction rows were found.');
    for (const g of groups.values()) {
      if (g.buffer) await g.stream.write(g.buffer);
      await g.stream.write(XML_END);
      await g.stream.close(); streams.delete(g.stream);
    }
    notify('Generating workbooks', `${groups.size} party files`, 67);
    const zipName = 'LGP_Console_Partywise.zip';
    const output = await dir.getFileHandle(zipName, { create: true });
    const outerStream = await output.createWritable(); streams.add(outerStream);
    const outer = new ZipWriter(new WritableStreamWriter(outerStream), { level: 0 });
    let i = 0, reconciled = 0;
    for (const g of groups.values()) {
      const wbHandle = await dir.getFileHandle(`workbook_${i}.xlsx`, { create: true });
      const wbStream = await wbHandle.createWritable(); streams.add(wbStream);
      const zip = new ZipWriter(new WritableStreamWriter(wbStream));
      for (const [name, xml] of Object.entries(STATIC)) await zip.add(name, new TextReader(xml));
      await zip.add('xl/worksheets/sheet1.xml', new BlobReader(await g.handle.getFile()));
      await zip.close(); streams.delete(wbStream);
      await outer.add(`LGP_Console_${g.name}.xlsx`, new BlobReader(await wbHandle.getFile()), { level: 0 });
      reconciled += g.rows;

      notify('Generating workbooks', `Created ${i} of ${groups.size}`, 67 + Math.round(i / groups.size * 32));
    }
    await outer.close(); streams.delete(outerStream);
    if (reconciled !== rows) throw new Error('Row reconciliation failed.');
    await reader.close();
    self.postMessage({ type: 'done', dir: dirName, zipName, rows, parties: groups.size, blankParties });
  } catch (error) {
    for (const stream of streams) { try { await stream.abort(); } catch {} }
    try { await reader.close(); } catch {}
    throw error;
  }
}
self.onmessage = async event => {
  try { await run(event.data?.file); }
  catch (error) {
    self.postMessage({ type: 'error', message: error?.stack || error?.message || String(error) });
  }
};
