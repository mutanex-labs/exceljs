'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const {spawnSync} = require('child_process');
const {
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
  writeStreamingSheets,
} = require('./helpers');

const dateFixturePath = path.join(
  repositoryRoot,
  'spec',
  'integration',
  'data',
  'dateIssue.xlsx'
);
const raceTestPath = path.join(
  repositoryRoot,
  'spec',
  'integration',
  'workbook-xlsx-reader.spec.js'
);
const mochaPath = path.join(
  repositoryRoot,
  'node_modules',
  'mocha',
  'bin',
  '_mocha'
);
const fixturesDirectory = path.join(__dirname, 'fixtures');
const fixtureHashes = {
  dateStylesSharedStrings:
    'ba84f5335abbe19a462a946c3961e5ff5739c5e53e93bf32115ac60ffe646844',
  richTextSharedStrings:
    '294643d324598b8cfe5b058ed8e007e6e1fafec528708af8dd402e08e6a9a347',
  workbookReaderOrderRace:
    '02b0710cb327c38b8dfac7a68827c51b2e505dc03eb23ff2eef6cdd6791cf7b8',
};

function sha256(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

function compact(result) {
  if (!Array.isArray(result.rows)) return result;
  return {
    status: result.status,
    error: result.error,
    worksheetNames: result.worksheets.map(sheet => sheet.name),
    rowCount: result.rows.length,
    firstRow: result.rows[0] || null,
    lastRow: result.rows[result.rows.length - 1] || null,
  };
}

function signature(result) {
  if (result.status === 'error') {
    return `error:${result.error.code || result.error.message}`;
  }
  if (result.status === 'timeout') return 'timeout';
  if (!Array.isArray(result.rows)) {
    return `${result.status}:${
      result.rowCount || result.rowsBeforeAbort || 0
    }:${JSON.stringify(result.lastRow || null)}`;
  }
  return `end:${JSON.stringify(result.rows.map(row => row.values))}`;
}

async function repeatProbe(runs, probe) {
  const counts = {};
  const samples = [];
  for (let run = 1; run <= runs; run += 1) {
    // Keep repeated readers uncontended so run-to-run transport observations stay comparable.
    // eslint-disable-next-line no-await-in-loop
    const result = await probe(run);
    const key = signature(result);
    counts[key] = (counts[key] || 0) + 1;
    samples.push({run, ...compact(result)});
  }
  return {runs, counts, samples};
}

async function buildStreamingWorkbook() {
  const rows = [['name', 'number']];
  for (let index = 1; index <= 20; index += 1) {
    rows.push([`value-${index}`, index]);
  }
  const buffer = await writeStreamingSheets([{name: 'Data', rows}], {
    useSharedStrings: true,
    useStyles: true,
  });
  return reorderArchive(buffer, 'original');
}

async function buildDocumentWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');
  workbook.creator = 'ExcelJS characterization';
  workbook.created = new Date('2020-01-02T03:04:05.000Z');
  workbook.modified = new Date('2020-01-02T03:04:05.000Z');
  worksheet.addRow(['name', 'number']);
  for (let index = 1; index <= 20; index += 1) {
    worksheet.addRow([`value-${index}`, index]);
  }
  const buffer = await workbook.xlsx.writeBuffer({
    useSharedStrings: true,
    useStyles: true,
  });
  return reorderArchive(buffer, 'original');
}

async function slowConsumer(buffer) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(
    chunkReadable(buffer, [7, 31, 13]),
    {
      worksheets: 'emit',
      sharedStrings: 'cache',
      styles: 'ignore',
      hyperlinks: 'ignore',
    }
  );
  const result = {
    status: 'end',
    rowCount: 0,
    worksheetNames: [],
    firstRow: null,
    lastRow: null,
  };
  for await (const worksheet of reader) {
    result.worksheetNames.push(worksheet.name);
    for await (const row of worksheet) {
      result.rowCount += 1;
      const values = row.values.slice(1);
      if (!result.firstRow) result.firstRow = values;
      result.lastRow = values;
      await new Promise(resolve => setTimeout(resolve, 1));
    }
  }
  return result;
}

async function abortAfterRows(buffer, abortAfterRowsCount = 5) {
  const input = delayedChunkReadable(buffer, 256, 2);
  input.on('error', () => {});
  const result = {status: 'pending', rowsBeforeAbort: 0, terminalError: null};
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    styles: 'ignore',
    hyperlinks: 'ignore',
  });

  const terminal = new Promise(resolve => {
    reader.once('end', () => {
      result.status = 'end';
      resolve(result);
    });
    reader.once('error', error => {
      result.status = 'error';
      result.terminalError = {
        name: error.name,
        code: error.code || null,
        message: error.message,
      };
      resolve(result);
    });
  });
  reader.on('worksheet', worksheet => {
    worksheet.on('row', () => {
      result.rowsBeforeAbort += 1;
      if (result.rowsBeforeAbort === abortAfterRowsCount) input.destroy();
    });
  });
  reader.read();

  return Promise.race([
    terminal,
    new Promise(resolve =>
      setTimeout(() => {
        result.status = 'no-terminal-event-before-timeout';
        resolve(result);
      }, 1500)
    ),
  ]);
}

function runUpstreamRaceSpec() {
  const args = [
    mochaPath,
    '--reporter',
    'json',
    '--require',
    path.join(repositoryRoot, 'spec', 'config', 'setup.js'),
    raceTestPath,
    '--grep',
    'WorkbookReader Serialise xlsx file',
  ];
  const execution = spawnSync(process.execPath, args, {
    cwd: repositoryRoot,
    encoding: 'utf8',
    timeout: 30000,
    windowsHide: true,
  });

  let report;
  try {
    report = JSON.parse(execution.stdout || '{}');
  } catch (error) {
    report = null;
  }
  return {
    exitStatus: execution.status,
    signal: execution.signal || null,
    timedOut: Boolean(execution.error && execution.error.code === 'ETIMEDOUT'),
    executionError: execution.error
      ? {code: execution.error.code, message: execution.error.message}
      : null,
    passed: report && report.stats ? report.stats.passes : 0,
    failed: report && report.stats ? report.stats.failures : 0,
    failures:
      report && Array.isArray(report.failures)
        ? report.failures.map(failure => ({
            title: failure.fullTitle,
            message: failure.err && failure.err.message,
          }))
        : [],
    stderr: execution.stderr || '',
  };
}

function assertFixtureHash(name, buffer, expected) {
  const actual = sha256(buffer);
  if (actual !== expected) {
    throw new Error(
      `${name} fixture hash changed: expected ${expected}, got ${actual}`
    );
  }
  return actual;
}

async function assertSameArchiveParts(
  leftBuffer,
  rightBuffer,
  description = 'reference fixture'
) {
  const [left, right] = await Promise.all([
    JSZip.loadAsync(leftBuffer),
    JSZip.loadAsync(rightBuffer),
  ]);
  const names = archive =>
    Object.keys(archive.files)
      .filter(name => !archive.files[name].dir)
      .sort();
  const leftNames = names(left);
  const rightNames = names(right);
  if (JSON.stringify(leftNames) !== JSON.stringify(rightNames)) {
    throw new Error(`Archive part names differ from ${description}`);
  }
  await Promise.all(
    leftNames.map(async name => {
      const [leftPart, rightPart] = await Promise.all([
        left.file(name).async('nodebuffer'),
        right.file(name).async('nodebuffer'),
      ]);
      if (!leftPart.equals(rightPart)) {
        throw new Error(
          `Archive part content differs from ${description}: ${name}`
        );
      }
    })
  );
  return {
    partNames: leftNames,
    zipEntryOrder: Object.keys(right.files).filter(
      name => !right.files[name].dir
    ),
  };
}

async function main() {
  ensureOutputDirectory();
  const runsArgument = process.argv.find(argument =>
    argument.startsWith('--runs=')
  );
  const raceRunsArgument = process.argv.find(argument =>
    argument.startsWith('--race-runs=')
  );
  const runs = Number((runsArgument && runsArgument.split('=')[1]) || 10);
  const raceRuns = Number(
    (raceRunsArgument && raceRunsArgument.split('=')[1]) || 20
  );
  const baseBuffer = await buildStreamingWorkbook();
  const documentBuffer = await buildDocumentWorkbook();
  const minimalDatePath = path.join(
    fixturesDirectory,
    'minimal-styles-date-sharedstrings.xlsx'
  );
  const minimalDateBuffer = fs.readFileSync(minimalDatePath);
  const generatedDateBuffer =
    await createMinimalStyledDateSharedStringWorkbook();
  assertFixtureHash(
    'Generated date/styles/sharedStrings',
    generatedDateBuffer,
    fixtureHashes.dateStylesSharedStrings
  );
  const richFixturePath = path.join(
    fixturesDirectory,
    'minimal-richtext-sharedstrings.xlsx'
  );
  const richBuffer = fs.readFileSync(richFixturePath);
  assertFixtureHash(
    'Rich-text/sharedStrings',
    richBuffer,
    fixtureHashes.richTextSharedStrings
  );
  const richZip = await JSZip.loadAsync(richBuffer);
  const richEntryOrder = Object.keys(richZip.files).filter(
    name => !richZip.files[name].dir
  );
  const richReaderReadyBuffer = await reorderArchive(
    richBuffer,
    'reader-ready-before-worksheet'
  );
  const richReaderReadyPath = path.join(
    outputDirectory,
    'rich-text-reader-ready-order.xlsx'
  );
  fs.writeFileSync(richReaderReadyPath, richReaderReadyBuffer);
  const richArchiveParts = await assertSameArchiveParts(
    richBuffer,
    richReaderReadyBuffer,
    'the frozen rich-text fixture'
  );
  const raceFixturePath = path.join(
    fixturesDirectory,
    'minimal-workbookreader-order-race.xlsx'
  );
  const raceArchive = fs.readFileSync(raceFixturePath);
  assertFixtureHash(
    'WorkbookReader ZIP-order race',
    raceArchive,
    fixtureHashes.workbookReaderOrderRace
  );
  const generatedRaceArchive = await reorderArchive(
    generatedDateBuffer,
    'worksheet-before-workbook'
  );
  assertFixtureHash(
    'Generated WorkbookReader ZIP-order race',
    generatedRaceArchive,
    fixtureHashes.workbookReaderOrderRace
  );
  const raceArchiveParts = await assertSameArchiveParts(
    minimalDateBuffer,
    raceArchive
  );
  const raceEntryOrder = raceArchiveParts.zipEntryOrder;
  const baseSha256 = sha256(baseBuffer);
  const patterns = [
    {name: 'chunk-1', pattern: 1},
    {name: 'chunk-7', pattern: 7},
    {name: 'chunk-64', pattern: 64},
    {name: 'chunk-1024', pattern: 1024},
    {name: 'irregular-1-7-64-3', pattern: [1, 7, 64, 3]},
  ];
  const chunkBoundaries = [];
  for (const item of patterns) {
    chunkBoundaries.push({
      name: item.name,
      category: 'stress/invariant probe',
      // eslint-disable-next-line no-await-in-loop
      result: await repeatProbe(runs, async () =>
        readStreaming(chunkReadable(baseBuffer, item.pattern), {
          timeoutMs: 10000,
        })
      ),
    });
  }

  const ordinaryWorkbookReader = {
    category: 'stress/invariant probe',
    generatedBy:
      'ExcelJS.Workbook.xlsx.writeBuffer; same cells as the WorkbookWriter fixture',
    inputSha256: sha256(documentBuffer),
    chunkBoundaries: [],
  };
  for (const item of [patterns[0], patterns[2], patterns[3]]) {
    ordinaryWorkbookReader.chunkBoundaries.push({
      name: item.name,
      // eslint-disable-next-line no-await-in-loop
      result: await repeatProbe(runs, async () =>
        readStreaming(chunkReadable(documentBuffer, item.pattern), {
          timeoutMs: 10000,
        })
      ),
    });
  }

  const zipOrderInputs = await Promise.all(
    [
      {name: 'original-entry-order', mode: 'original'},
      {name: 'reverse-entry-order', mode: 'reverse'},
      {
        name: 'shared-strings-and-relationships-then-worksheet-before-workbook',
        mode: 'worksheet-before-workbook',
      },
    ].map(async item => ({
      ...item,
      archive:
        item.mode === 'original'
          ? baseBuffer
          : await reorderArchive(baseBuffer, item.mode),
    }))
  );
  const zipOrders = [];
  for (const item of zipOrderInputs) {
    const {archive} = item;
    zipOrders.push({
      name: item.name,
      category:
        item.mode === 'worksheet-before-workbook'
          ? 'known-upstream-bug'
          : 'stress/invariant probe',
      inputSha256: sha256(archive),
      // eslint-disable-next-line no-await-in-loop
      result: await repeatProbe(runs, async () =>
        readStreaming(chunkReadable(archive, 64), {timeoutMs: 10000})
      ),
    });
  }

  const dateBuffer = fs.readFileSync(dateFixturePath);
  const dateResult = await repeatProbe(runs, async () =>
    readStreaming(chunkReadable(dateBuffer, dateBuffer.length), {
      timeoutMs: 10000,
      readerOptions: {
        styles: 'cache',
        sharedStrings: 'cache',
        hyperlinks: 'ignore',
      },
    })
  );
  const minimalDateResult = await repeatProbe(runs, async () =>
    readStreaming(chunkReadable(minimalDateBuffer, minimalDateBuffer.length), {
      timeoutMs: 10000,
      readerOptions: {
        styles: 'cache',
        sharedStrings: 'cache',
        hyperlinks: 'ignore',
      },
    })
  );

  const richTextResult = await repeatProbe(runs, async () =>
    readStreaming(chunkReadable(richBuffer, richBuffer.length), {
      timeoutMs: 10000,
    })
  );
  const richTextReaderReadyResult = await repeatProbe(runs, async () =>
    readStreaming(
      chunkReadable(richReaderReadyBuffer, richReaderReadyBuffer.length),
      {timeoutMs: 10000}
    )
  );
  const richTextChunkBoundaries = [];
  for (const item of patterns) {
    richTextChunkBoundaries.push({
      name: item.name,
      category: 'known-upstream-bug',
      // eslint-disable-next-line no-await-in-loop
      result: await repeatProbe(runs, async () =>
        readStreaming(chunkReadable(richBuffer, item.pattern), {
          timeoutMs: 10000,
        })
      ),
    });
  }

  const slowConsumerResult = await repeatProbe(Math.min(runs, 5), async () =>
    slowConsumer(baseBuffer)
  );
  const truncatedResult = await repeatProbe(runs, async () => {
    const truncated = baseBuffer.subarray(
      0,
      Math.floor(baseBuffer.length * 0.6)
    );
    return readStreaming(
      chunkReadable(truncated, Math.max(1, truncated.length)),
      {timeoutMs: 10000}
    );
  });

  const abortBuffer = await writeStreamingSheets(
    [{name: 'Abort', rows: 5000}],
    {useSharedStrings: true, useStyles: false}
  );
  const abortArchive = await reorderArchive(
    abortBuffer,
    'reader-ready-before-worksheet'
  );
  const abortResult = await repeatProbe(Math.min(runs, 5), async () =>
    abortAfterRows(abortArchive)
  );

  const deterministicRace = await repeatProbe(raceRuns, async () =>
    readStreaming(chunkReadable(raceArchive, 64), {timeoutMs: 10000})
  );
  const upstreamRaceRuns = [];
  for (let run = 1; run <= raceRuns; run += 1) {
    upstreamRaceRuns.push({run, ...runUpstreamRaceSpec()});
  }

  const report = {
    baselineCommit: '5bed18b45e824f409b08456b59b87430ded023ab',
    createdAt: new Date().toISOString(),
    environment: {
      node: process.version,
      platform: process.platform,
      architecture: process.arch,
    },
    fixture: {
      logicalWorkbookSha256: baseSha256,
      generatedBy:
        'ExcelJS WorkbookWriter with fixed metadata and normalized ZIP timestamps, 1 header row and 20 data rows; useSharedStrings and useStyles enabled',
      upstreamIssue1328: {
        path: path.relative(repositoryRoot, dateFixturePath),
        sha256: sha256(dateBuffer),
        provenance: 'Existing ExcelJS upstream integration fixture, 4.7 KB',
        category: 'known-upstream-bug',
        result: dateResult,
        firstRowValueKinds: dateResult.samples.map(sample => {
          const counts = {};
          for (const value of (sample.firstRow && sample.firstRow.values) ||
            []) {
            const key =
              value.type === 'sharedString'
                ? `sharedString:${value.index}`
                : value.type;
            counts[key] = (counts[key] || 0) + 1;
          }
          return counts;
        }),
      },
      minimalStylesDateSharedStrings: {
        path: path.relative(repositoryRoot, minimalDatePath),
        bytes: minimalDateBuffer.length,
        sha256: sha256(minimalDateBuffer),
        provenance:
          'Programmatically generated seven-part OOXML package; one shared-string header plus one styled Excel date; sharedStrings entry follows worksheet entry in ZIP order',
        category: 'known-upstream-bug',
        result: minimalDateResult,
      },
      pullRequest1431: {
        path: path.relative(repositoryRoot, richFixturePath),
        bytes: richBuffer.length,
        sha256: sha256(richBuffer),
        zipEntryOrder: richEntryOrder,
        provenance:
          'Frozen single-row WorkbookWriter fixture from the baseline, one rich shared string and one plain shared string',
        category: 'known-upstream-bug',
        result: richTextResult,
        readerReadyOrder: {
          path: path.relative(repositoryRoot, richReaderReadyPath),
          sha256: sha256(richReaderReadyBuffer),
          samePartContent: true,
          partNames: richArchiveParts.partNames,
          zipEntryOrder: richArchiveParts.zipEntryOrder,
          result: richTextReaderReadyResult,
        },
        chunkBoundaries: richTextChunkBoundaries,
      },
      validZipOrderRace: {
        path: path.relative(repositoryRoot, raceFixturePath),
        bytes: raceArchive.length,
        sha256: sha256(raceArchive),
        logicalFixtureSha256: sha256(minimalDateBuffer),
        samePartContentAsDateFixture: true,
        partNames: raceArchiveParts.partNames,
        zipEntryOrder: raceEntryOrder,
        provenance:
          'Same minimal generated styles/date/sharedStrings OOXML package with only ZIP entry order changed; relationships and sharedStrings precede worksheet, workbook.xml follows worksheet',
        category: 'known-upstream-bug',
        expectedFailure:
          'TypeError reading sheets from undefined in WorkbookReader._parseWorksheet',
        result: deterministicRace,
      },
    },
    chunkBoundaries,
    ordinaryWorkbookReader,
    zipOrders,
    slowConsumerBackpressure: {
      category: 'stress/invariant probe',
      result: slowConsumerResult,
    },
    earlyAbort: {
      category: 'stress/invariant probe',
      inputRows: 5000,
      abortAfterRows: 5,
      result: abortResult,
    },
    truncatedInput: {
      category: 'stress/invariant probe',
      truncationRatio: 0.6,
      result: truncatedResult,
    },
    upstreamFlakyRace: {
      category: 'known-upstream-bug',
      test: 'WorkbookReader Serialise xlsx file',
      fixture:
        'Existing upstream test workload generated by testUtils.createTestBook(..., "xlsx")',
      runs: raceRuns,
      outcomes: upstreamRaceRuns,
      failureCount: upstreamRaceRuns.filter(
        item => item.failed > 0 || item.exitStatus !== 0
      ).length,
      passCount: upstreamRaceRuns.filter(
        item => item.passed > 0 && item.failed === 0
      ).length,
    },
  };

  const outputPath = path.join(outputDirectory, 'streaming-adversarial.json');
  fs.writeFileSync(outputPath, `${JSON.stringify(report, null, 2)}\n`);
  process.stdout.write(
    `${JSON.stringify(
      {
        outputPath,
        raceFailures: report.upstreamFlakyRace.failureCount,
        racePasses: report.upstreamFlakyRace.passCount,
      },
      null,
      2
    )}\n`
  );
}

main().catch(error => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
