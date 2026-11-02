
import {
  ZipReader,
  ZipWriter,
  BlobReader,
  TextReader,
  TextWriter,
  Writer
} from '@zip.js/zip.js';

class WritableStreamWriter extends Writer {
  constructor(stream) {
    super();
    this.writer = stream.getWriter();
    this.closed = false;
  }

  async writeUint8Array(array) {
    if (this.closed) {
      throw new Error('Cannot write to a closed stream.');
    }
    await this.writer.write(array);
  }

  async close() {
    if (this.closed) return;
    await this.writer.close();
    this.closed = true;
    this.writer.releaseLock();
  }
}

const XML_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';

const XML_END = '</sheetData></worksheet>';

const STATIC = {
  '[Content_Types].xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>',

  '_rels/.rels':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>',

  'xl/workbook.xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="LGP Console" sheetId="1" r:id="rId1"/></sheets>' +
    '</workbook>',

  'xl/_rels/workbook.xml.rels':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '</Relationships>'
};

const notify = (label, detail, percent) => {
  self.postMessage({ type: 'progress', label, detail, percent });
};

function decodeXml(value) {
  return String(value ?? '')
    .replace(/&#x([0-9a-f]+);/gi, (_, n) =>
      String.fromCodePoint(parseInt(n, 16))
    )
    .replace(/&#(\d+);/g, (_, n) =>
      String.fromCodePoint(Number(n))
    )
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function normalize(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function safeName(value) {
  return String(value)
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 85) || 'UNASSIGNED';
}

function columnIndex(letters) {
  let result = 0;
  for (const char of letters) {
    result = result * 26 + char.charCodeAt(0) - 64;
  }
  return result - 1;
}

function cellParts(row) {
  const cells = [];
  const regex = /<c\b([^>]*?)(?:\s*\/>|>([\s\S]*?)<\/c\s*>)/g;
  let match;

  while ((match = regex.exec(row))) {
    const reference = /\br="([A-Z]+)(\d+)"/.exec(match[1]);
    if (reference) {
      cells.push({
        index: columnIndex(reference[1]),
        col: reference[1],
        attrs: match[1],
        body: match[2] || ''
      });
    }
  }
  return cells;
}

function extractCell(cell, sharedStrings) {
  const type = /\bt="([^"]+)"/.exec(cell.attrs)?.[1] || '';
  const raw = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(cell.body)?.[1];

  if (type === 's') {
    const index = Number(raw);
    if (
      !Number.isInteger(index) ||
      index < 0 ||
      index >= sharedStrings.length
    ) {
      throw new Error(`Invalid shared string index: ${raw}`);
    }
    return sharedStrings[index];
  }

  if (type === 'inlineStr') {
    return [...cell.body.matchAll(
      /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g
    )].map(match => decodeXml(match[1])).join('');
  }

  return decodeXml(raw ?? '');
}

function cleanRow(row, rowNumber, sharedStrings) {
  let xml = `<row r="${rowNumber}">`;

  for (const cell of cellParts(row)) {
    const type = /\bt="([^"]+)"/.exec(cell.attrs)?.[1] || '';
    const value = extractCell(cell, sharedStrings);
    const reference = `${cell.col}${rowNumber}`;

    if (['s', 'inlineStr', 'str', 'e'].includes(type)) {
      xml +=
        `<c r="${reference}" t="inlineStr">` +
        `<is><t xml:space="preserve">${escapeXml(value)}</t></is>` +
        '</c>';
    } else if (type === 'b') {
      xml +=
        `<c r="${reference}" t="b">` +
        `<v>${value === '1' ? '1' : '0'}</v></c>`;
    } else if (value !== '') {
      if (!Number.isFinite(Number(value))) {
        throw new Error(`Unexpected non-numeric value in ${reference}`);
      }
      xml +=
        `<c r="${reference}" t="n"><v>${value}</v></c>`;
    }
  }

  return xml + '</row>';
}

function parseSharedStrings(xml) {
  const values = [];

  for (const item of xml.matchAll(
    /<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g
  )) {
    values.push(
      [...item[1].matchAll(
        /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g
      )].map(match => decodeXml(match[1])).join('')
    );
  }
  return values;
}

async function run(file) {
  if (!file) throw new Error('Select an XLSX workbook.');

  if (!self.navigator?.storage?.getDirectory) {
    throw new Error(
      'Browser storage is unavailable. Use Chrome or Edge over HTTPS.'
    );
  }

  const root = await self.navigator.storage.getDirectory();
  const dirName = `lgp-split-${Date.now()}`;
  const dir = await root.getDirectoryHandle(dirName, { create: true });
  const reader = new ZipReader(new BlobReader(file));
  const openStreams = new Set();

  try {
    const entries = await reader.getEntries();
    const sheet =
      entries.find(e => e.filename === 'xl/worksheets/sheet1.xml') ||
      entries.find(e =>
        /^xl\/worksheets\/sheet\d+\.xml$/.test(e.filename)
      );

    if (!sheet) throw new Error('No worksheet found in this XLSX file.');

    const sharedEntry = entries.find(
      e => e.filename === 'xl/sharedStrings.xml'
    );
    let sharedStrings = [];

    if (sharedEntry) {
      notify('Reading shared strings', 'Preparing Excel text values', 2);
      const text = await sharedEntry.getData(new TextWriter());
      sharedStrings = parseSharedStrings(text);
    }

    const groups = new Map();
    const usedNames = new Set();
    let headers = null;
    let partyIndex = -1;
    let headerXml = '';
    let rows = 0;
    let blankParties = 0;

    async function getGroup(party) {
      const key = normalize(party) || '__UNASSIGNED__';
      if (groups.has(key)) return groups.get(key);

      let name = safeName(
        key === '__UNASSIGNED__' ? 'UNASSIGNED' : key
      );
      const base = name;
      let suffix = 2;

      while (usedNames.has(name.toLowerCase())) {
        name = `${base.slice(0, 75)}_${suffix++}`;
      }
      usedNames.add(name.toLowerCase());

      const handle = await dir.getFileHandle(
        `party_${groups.size}.xml`,
        { create: true }
      );
      const stream = await handle.createWritable();
      openStreams.add(stream);

      await stream.write(XML_HEAD + headerXml);

      const group = {
        key,
        name,
        handle,
        stream,
        rows: 0,
        buffer: ''
      };
      groups.set(key, group);
      return group;
    }

    async function acceptRow(row) {
      if (!headers) {
        headers = [];
        for (const cell of cellParts(row)) {
          headers[cell.index] = extractCell(cell, sharedStrings);
        }

        partyIndex = headers.findIndex(
          h => normalize(h).replace(/\s+/g, '_') === 'PARTY_NAME'
        );
        if (partyIndex < 0) {
          throw new Error(
            'Required column "Party Name" or "party_name" was not found.'
          );
        }

        headerXml = cleanRow(row, 1, sharedStrings);
        return;
      }

      const partyCell = cellParts(row).find(
        cell => cell.index === partyIndex
      );
      const party = partyCell
        ? extractCell(partyCell, sharedStrings)
        : '';

      if (!String(party).trim()) blankParties++;

      const group = await getGroup(party);
      group.rows++;
      rows++;
      group.buffer += cleanRow(row, group.rows + 1, sharedStrings);

      if (group.buffer.length >= 256 * 1024) {
        await group.stream.write(group.buffer);
        group.buffer = '';
      }

      if (rows % 10000 === 0) {
        notify(
          'Grouping transactions',
          `${rows.toLocaleString()} rows, ${groups.size} parties`,
          Math.min(65, 5 + Math.round(rows / 450000 * 60))
        );
      }
    }

    notify('Reading workbook', 'Streaming worksheet data', 4);

    let pending = '';
    let inData = false;
    const decoder = new TextDecoder();

    async function parseChunk(chunk) {
      pending += decoder.decode(chunk, { stream: true });

      if (!inData) {
        const match = /<sheetData(?:\s[^>]*)?>/.exec(pending);
        if (!match) {
          if (pending.length > 2_000_000) {
            throw new Error('Could not locate worksheet data.');
          }
          return;
        }
        pending = pending.slice(match.index + match[0].length);
        inData = true;
      }

      while (true) {
        const start = /<row(?:\s|>)/.exec(pending);
        if (!start) {
          pending = pending.slice(-128);
          return;
        }
        if (start.index) pending = pending.slice(start.index);

        const end = pending.indexOf('</row>');
        if (end < 0) {
          if (pending.length > 20_000_000) {
            throw new Error('A worksheet row is unexpectedly large.');
          }
          return;
        }

        const row = pending.slice(0, end + 6);
        pending = pending.slice(end + 6);
        await acceptRow(row);
      }
    }

    await sheet.getData(
      new WritableStreamWriter(
        new WritableStream({ write: parseChunk })
      )
    );

    if (!headers || rows === 0) {
      throw new Error('No transaction rows were found.');
    }

    // Finish writing each party's worksheet XML.
    for (const group of groups.values()) {
      if (group.buffer) await group.stream.write(group.buffer);
      await group.stream.write(XML_END);
      await group.stream.close();
      openStreams.delete(group.stream);
    }

    notify(
      'Generating workbooks',
      `${groups.size} party files`,
      67
    );

    const zipName = 'LGP_Console_Partywise.zip';
    const outputHandle = await dir.getFileHandle(
      zipName,
      { create: true }
    );
    const outerStream = await outputHandle.createWritable();
    openStreams.add(outerStream);

    const outerZip = new ZipWriter(
      new WritableStreamWriter(outerStream),
      { level: 0 }
    );

    let generated = 0;
    let reconciledRows = 0;

    for (const group of groups.values()) {
      const workbookHandle = await dir.getFileHandle(
        `workbook_${generated}.xlsx`,
        { create: true }
      );
      const workbookStream = await workbookHandle.createWritable();
      openStreams.add(workbookStream);

      const workbookZip = new ZipWriter(
        new WritableStreamWriter(workbookStream)
      );

      for (const [name, xml] of Object.entries(STATIC)) {
        await workbookZip.add(name, new TextReader(xml));
      }

      await workbookZip.add(
        'xl/worksheets/sheet1.xml',
        new BlobReader(await group.handle.getFile())
      );

      await workbookZip.close();
      openStreams.delete(workbookStream);

      const workbookFile = await workbookHandle.getFile();
      if (workbookFile.size === 0) {
        throw new Error(
          `Generated workbook ${group.name} is empty.`
        );
      }

      await outerZip.add(
        `LGP_Console_${group.name}.xlsx`,
        new BlobReader(workbookFile),
        { level: 0 }
      );

      reconciledRows += group.rows;
      generated++;

      notify(
        'Generating workbooks',
        `Created ${generated} of ${groups.size}`,
        67 + Math.round(generated / groups.size * 32)
      );
    }

    await outerZip.close();
    openStreams.delete(outerStream);

    // Validate that the ZIP is committed and contains all files.
    const finalZipFile = await outputHandle.getFile();

    if (finalZipFile.size === 0) {
      throw new Error('ZIP generation failed: output file is empty.');
    }

    const verifyReader = new ZipReader(
      new BlobReader(finalZipFile)
    );

    try {
      const zipEntries = await verifyReader.getEntries();
      const xlsxEntries = zipEntries.filter(
        entry => entry.filename.toLowerCase().endsWith('.xlsx')
      );

      if (xlsxEntries.length !== groups.size) {
        throw new Error(
          `ZIP validation failed. Expected ${groups.size} Excel files, ` +
          `but found ${xlsxEntries.length}.`
        );
      }
    } finally {
      await verifyReader.close();
    }

    if (reconciledRows !== rows) {
      throw new Error('Transaction row reconciliation failed.');
    }

    await reader.close();

    notify(
      'Completed',
      `Generated ${groups.size} party-wise workbooks`,
      100
    );

    self.postMessage({
      type: 'done',
      dir: dirName,
      zipName,
      rows,
      parties: groups.size,
      blankParties
    });

  } catch (error) {
    for (const stream of openStreams) {
      try { await stream.abort(); } catch {}
    }
    try { await reader.close(); } catch {}
    throw error;
  }
}

self.onmessage = async event => {
  try {
    await run(event.data?.file);
  } catch (error) {
    self.postMessage({
      type: 'error',
      message: error?.stack || error?.message || String(error)
    });
  }
};
