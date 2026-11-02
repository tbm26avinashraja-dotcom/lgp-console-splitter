
import {
  ZipReader,
  ZipReaderStream,
  ZipWriterStream,
  BlobReader,
  TextWriter
} from '@zip.js/zip.js';

const XML_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData>';

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

const encoder = new TextEncoder();

function notify(label, detail, percent) {
  self.postMessage({
    type: 'progress',
    label,
    detail,
    percent
  });
}

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
    )]
      .map(match => decodeXml(match[1]))
      .join('');
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
        throw new Error(
          `Unexpected non-numeric value in ${reference}`
        );
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
      )]
        .map(match => decodeXml(match[1]))
        .join('')
    );
  }

  return values;
}

async function writeTextEntry(zip, name, text) {
  const entry = new WritableStream();
  const writer = entry.getWriter();
  await writer.write(encoder.encode(text));
  await writer.close();
}

async function addZipEntry(zipWriter, name, source) {
  const entry = new WritableStream();
  const writer = entry.getWriter();

  // Native zip.js stream API.
  const zipEntry = new ZipWriterStream({
    level: 0
  });

  const sourceStream = source.stream();

  await sourceStream.pipeTo(
    zipEntry.writable
  );

  return zipEntry.readable;
}

async function run(file) {
  if (!file) {
    throw new Error('Select an XLSX workbook.');
  }

  if (!self.navigator?.storage?.getDirectory) {
    throw new Error(
      'Browser storage is unavailable. Use Chrome or Edge over HTTPS.'
    );
  }

  const root = await self.navigator.storage.getDirectory();

  const dirName = `lgp-split-${Date.now()}`;

  const dir = await root.getDirectoryHandle(dirName, {
    create: true
  });

  const inputReader = new ZipReader(new BlobReader(file));

  try {
    const entries = await inputReader.getEntries();

    const sheet =
      entries.find(
        entry => entry.filename === 'xl/worksheets/sheet1.xml'
      ) ||
      entries.find(
        entry => /^xl\/worksheets\/sheet\d+\.xml$/.test(
          entry.filename
        )
      );

    if (!sheet) {
      throw new Error('No worksheet found in this XLSX file.');
    }

    const sharedEntry = entries.find(
      entry => entry.filename === 'xl/sharedStrings.xml'
    );

    let sharedStrings = [];

    if (sharedEntry) {
      notify(
        'Reading shared strings',
        'Preparing Excel text values',
        2
      );

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

      if (groups.has(key)) {
        return groups.get(key);
      }

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
          headers[cell.index] = extractCell(
            cell,
            sharedStrings
          );
        }

        partyIndex = headers.findIndex(
          header =>
            normalize(header).replace(/\s+/g, '_') ===
            'PARTY_NAME'
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

      if (!String(party).trim()) {
        blankParties++;
      }

      const group = await getGroup(party);

      group.rows++;
      rows++;

      group.buffer += cleanRow(
        row,
        group.rows + 1,
        sharedStrings
      );

      if (group.buffer.length >= 256 * 1024) {
        await group.stream.write(group.buffer);
        group.buffer = '';
      }

      if (rows % 10000 === 0) {
        notify(
          'Grouping transactions',
          `${rows.toLocaleString()} rows, ${groups.size} parties`,
          Math.min(
            65,
            5 + Math.round(rows / 450000 * 60)
          )
        );
      }
    }

    notify(
      'Reading workbook',
      'Streaming worksheet data',
      4
    );

    let pending = '';
    let inData = false;

    const decoder = new TextDecoder();

    async function parseChunk(chunk) {
      pending += decoder.decode(chunk, {
        stream: true
      });

      if (!inData) {
        const match = /<sheetData(?:\s[^>]*)?>/.exec(pending);

        if (!match) {
          if (pending.length > 2_000_000) {
            throw new Error(
              'Could not locate worksheet data.'
            );
          }

          return;
        }

        pending = pending.slice(
          match.index + match[0].length
        );

        inData = true;
      }

      while (true) {
        const start = /<row(?:\s|>)/.exec(pending);

        if (!start) {
          pending = pending.slice(-128);
          return;
        }

        if (start.index) {
          pending = pending.slice(start.index);
        }

        const end = pending.indexOf('</row>');

        if (end < 0) {
          if (pending.length > 20_000_000) {
            throw new Error(
              'A worksheet row is unexpectedly large.'
            );
          }

          return;
        }

        const row = pending.slice(0, end + 6);
        pending = pending.slice(end + 6);

        await acceptRow(row);
      }
    }

    // Stream worksheet XML from the source ZIP.
    const sourceStream = new WritableStream({
      write: parseChunk
    });

    await sheet.getData({
      writeUint8Array: async chunk => {
        await parseChunk(chunk);
      }
    });

    if (!headers || rows === 0) {
      throw new Error(
        'No transaction rows were found.'
      );
    }

    for (const group of groups.values()) {
      if (group.buffer) {
        await group.stream.write(group.buffer);
      }

      await group.stream.write(XML_END);
      await group.stream.close();
    }

    notify(
      'Generating workbooks',
      `${groups.size} party files`,
      67
    );

    // The output ZIP is generated using zip.js streams.
    const zipName = 'LGP_Console_Partywise.zip';

    const outputHandle = await dir.getFileHandle(
      zipName,
      { create: true }
    );

    const outputStream = await outputHandle.createWritable();

    const zipStream = new ZipWriterStream({
      level: 0
    });

    const zipOutputPromise = zipStream.readable.pipeTo(
      outputStream
    );

    let generated = 0;
    let reconciledRows = 0;

    for (const group of groups.values()) {
      // Generate a workbook for this Party Name.
      const workbookHandle = await dir.getFileHandle(
        `workbook_${generated}.xlsx`,
        { create: true }
      );

      const workbookOutput =
        await workbookHandle.createWritable();

      const workbookZip = new ZipWriterStream();

      const workbookPromise =
        workbookZip.readable.pipeTo(workbookOutput);

      // Add the required XLSX components.
      for (const [name, xml] of Object.entries(STATIC)) {
        const entry = new Blob([xml]).stream();

        await entry.pipeTo(
          workbookZip.writable,
          { preventClose: true }
        );
      }

      const worksheetFile =
        await group.handle.getFile();

      await worksheetFile.stream().pipeTo(
        workbookZip.writable,
        { preventClose: true }
      );

      await workbookZip.writable.close();
      await workbookPromise;

      const workbookFile =
        await workbookHandle.getFile();

      if (workbookFile.size === 0) {
        throw new Error(
          `Generated workbook ${group.name} is empty.`
        );
      }

      await workbookFile.stream().pipeTo(
        zipStream.writable,
        { preventClose: true }
      );

      reconciledRows += group.rows;
      generated++;

      notify(
        'Generating workbooks',
        `Created ${generated} of ${groups.size}`,
        67 + Math.round(
          generated / groups.size * 32
        )
      );
    }

    await zipStream.writable.close();
    await zipOutputPromise;

    const finalZipFile = await outputHandle.getFile();

    if (finalZipFile.size === 0) {
      throw new Error(
        'ZIP generation failed: output file is empty.'
      );
    }

    if (reconciledRows !== rows) {
      throw new Error(
        'Transaction row reconciliation failed.'
      );
    }

    await inputReader.close();

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
    try {
      await inputReader.close();
    } catch {}

    throw error;
  }
}

self.onmessage = async event => {
  try {
    await run(event.data?.file);
  } catch (error) {
    self.postMessage({
      type: 'error',
      message:
        error?.stack ||
        error?.message ||
        String(error)
    });
  }
};
