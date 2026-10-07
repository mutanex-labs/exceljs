'use strict';

const fs = require('fs');
const path = require('path');
const {PassThrough, Readable} = require('stream');
const JSZip = require('jszip');
const ExcelJS = require('../..');

const repositoryRoot = path.resolve(__dirname, '..', '..');
const outputDirectory = path.join(repositoryRoot, '.temp', 'characterization');

function ensureOutputDirectory() {
  fs.mkdirSync(outputDirectory, {recursive: true});
}

function chunkReadable(buffer, pattern) {
  const chunkSizes = Array.isArray(pattern) ? pattern : [pattern];
  let offset = 0;
  let index = 0;
  return new Readable({
    read() {
      if (offset >= buffer.length) {
        this.push(null);
        return;
      }
      const size = chunkSizes[index % chunkSizes.length];
      const nextOffset = Math.min(offset + size, buffer.length);
      this.push(buffer.subarray(offset, nextOffset));
      offset = nextOffset;
      index += 1;
    },
  });
}

function delayedChunkReadable(buffer, size, delayMs) {
  let offset = 0;
  let reading = false;
  return new Readable({
    read() {
      if (reading) return;
      reading = true;
      setTimeout(() => {
        reading = false;
        if (this.destroyed) return;
        if (offset >= buffer.length) {
          this.push(null);
          return;
        }
        const nextOffset = Math.min(offset + size, buffer.length);
        this.push(buffer.subarray(offset, nextOffset));
        offset = nextOffset;
      }, delayMs);
    },
  });
}

async function writeStreamingSheets(sheets, options = {}) {
  const output = new PassThrough();
  const chunks = [];
  output.on('data', chunk => chunks.push(chunk));
  const outputFinished = new Promise((resolve, reject) => {
    output.once('finish', resolve);
    output.once('error', reject);
  });

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: output,
    useSharedStrings: options.useSharedStrings !== false,
    useStyles: options.useStyles !== false,
    creator: options.creator || 'ExcelJS characterization',
    created: options.created || new Date('2020-01-02T03:04:05.000Z'),
    modified: options.modified || new Date('2020-01-02T03:04:05.000Z'),
  });
  for (const sheet of sheets) {
    const worksheet = workbook.addWorksheet(sheet.name);
    const rowCount =
      typeof sheet.rows === 'number' ? sheet.rows : sheet.rows.length;
    for (let index = 0; index < rowCount; index += 1) {
      const row =
        typeof sheet.rows === 'number'
          ? [index + 1, `row-${index + 1}`]
          : sheet.rows[index];
      worksheet.addRow(row);
    }
  }

  await Promise.all([workbook.commit(), outputFinished]);
  return Buffer.concat(chunks);
}

function writeStreamingWorkbook(rows, options = {}) {
  return writeStreamingSheets(
    [{name: options.sheetName || 'Data', rows}],
    options
  );
}

async function createMinimalStyledDateSharedStringWorkbook() {
  const packageRelationships =
    'http://schemas.openxmlformats.org/package/2006/relationships';
  const spreadsheet =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const zip = new JSZip();
  const timestamp = '2020-01-01T00:00:00.000Z';
  const parts = [
    [
      '[Content_Types].xml',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    ],
    [
      '_rels/.rels',
      `<Relationships xmlns="${packageRelationships}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ],
    [
      'xl/_rels/workbook.xml.rels',
      `<Relationships xmlns="${packageRelationships}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    ],
    [
      'xl/workbook.xml',
      `<workbook xmlns="${spreadsheet}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="false"/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ],
    [
      'xl/styles.xml',
      `<styleSheet xmlns="${spreadsheet}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
    ],
    [
      'xl/worksheets/sheet1.xml',
      `<worksheet xmlns="${spreadsheet}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" s="1" t="n"><v>44155</v></c></row></sheetData></worksheet>`,
    ],
    [
      'xl/sharedStrings.xml',
      `<sst xmlns="${spreadsheet}" count="1" uniqueCount="1"><si><t>Date</t></si></sst>`,
    ],
  ];

  // This is valid OOXML part ordering but deliberately puts sharedStrings after its worksheet.
  for (const [name, content] of parts) {
    zip.file(name, content, {date: new Date(timestamp), createFolders: false});
  }
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: {level: 9},
  });
}

function summarizeValue(value) {
  if (value instanceof Date) {
    return {type: 'date', value: value.toISOString()};
  }
  if (value && typeof value === 'object') {
    if (Object.prototype.hasOwnProperty.call(value, 'sharedString')) {
      return {type: 'sharedString', index: value.sharedString};
    }
    if (Array.isArray(value.richText)) {
      return {type: 'richText', runs: value.richText};
    }
    if (Object.prototype.hasOwnProperty.call(value, 'formula')) {
      return {
        type: 'formula',
        formula: value.formula,
        result: summarizeValue(value.result),
      };
    }
  }
  return {type: value === null ? 'null' : typeof value, value};
}

function summarizeRow(row) {
  return row.values.slice(1).map(summarizeValue);
}

function readStreaming(input, options = {}) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    styles: 'cache',
    hyperlinks: 'cache',
    entries: 'ignore',
    ...options.readerOptions,
  });
  const result = {worksheets: [], rows: [], status: 'pending', error: null};

  return new Promise(resolve => {
    let settled = false;
    const timeout = setTimeout(() => {
      result.status = 'timeout';
      finish();
    }, options.timeoutMs || 5000);

    function finish(status) {
      if (settled) return;
      settled = true;
      if (status) result.status = status;
      clearTimeout(timeout);
      resolve(result);
    }

    reader.on('worksheet', worksheet => {
      result.worksheets.push({id: worksheet.id, name: worksheet.name});
      worksheet.on('row', row => {
        result.rows.push({
          worksheet: worksheet.name || worksheet.id,
          number: row.number,
          values: summarizeRow(row),
        });
        if (typeof options.onRow === 'function') options.onRow(row, reader);
      });
    });
    reader.once('end', () => finish('end'));
    reader.once('error', error => {
      result.status = 'error';
      result.error = {
        name: error.name,
        code: error.code || null,
        message: error.message,
      };
      finish();
    });

    reader.read();
  });
}

async function reorderArchive(buffer, mode) {
  const source = await JSZip.loadAsync(buffer);
  const entries = Object.keys(source.files).filter(
    name => !source.files[name].dir
  );
  const originalIndex = new Map(entries.map((name, index) => [name, index]));
  const ordered = entries.slice();

  if (mode === 'reverse') {
    ordered.reverse();
  } else if (mode === 'worksheet-before-workbook') {
    const rank = name => {
      if (name === 'xl/sharedStrings.xml') return 0;
      if (name === 'xl/_rels/workbook.xml.rels') return 1;
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) return 2;
      if (name === 'xl/workbook.xml') return 3;
      return 4;
    };
    ordered.sort(
      (left, right) =>
        rank(left) - rank(right) ||
        originalIndex.get(left) - originalIndex.get(right)
    );
  } else if (mode === 'reader-ready-before-worksheet') {
    const rank = name => {
      if (name === 'xl/sharedStrings.xml') return 0;
      if (name === 'xl/_rels/workbook.xml.rels') return 1;
      if (name === 'xl/workbook.xml') return 2;
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) return 3;
      return 4;
    };
    ordered.sort(
      (left, right) =>
        rank(left) - rank(right) ||
        originalIndex.get(left) - originalIndex.get(right)
    );
  }

  const destination = new JSZip();
  const contents = await Promise.all(
    ordered.map(name => source.file(name).async('nodebuffer'))
  );
  contents.forEach((content, index) => {
    destination.file(ordered[index], content, {
      date: new Date('2020-01-01T00:00:00.000Z'),
      createFolders: false,
    });
  });
  return destination.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });
}

module.exports = {
  ExcelJS,
  JSZip,
  chunkReadable,
  createMinimalStyledDateSharedStringWorkbook,
  delayedChunkReadable,
  ensureOutputDirectory,
  outputDirectory,
  readStreaming,
  repositoryRoot,
  reorderArchive,
  summarizeRow,
  summarizeValue,
  writeStreamingSheets,
  writeStreamingWorkbook,
};
