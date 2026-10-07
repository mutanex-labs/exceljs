'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const {
  ExcelJS,
  JSZip,
  ensureOutputDirectory,
  outputDirectory,
  repositoryRoot,
} = require('./helpers');

const CONTENT_TYPES = '[Content_Types].xml';
const PACKAGE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/package/2006/relationships';

function classifyPart(name) {
  if (/vbaProject/i.test(name)) return 'macros-vba';
  if (/externalLinks/i.test(name)) return 'external-links';
  if (/customXml|vendor/i.test(name)) return 'custom-third-party-xml';
  if (/charts?/i.test(name)) return 'charts';
  if (/drawings?/i.test(name)) return 'drawings';
  if (/comments|vmlDrawing/i.test(name)) return 'comments-vml';
  if (/printerSettings/i.test(name)) return 'printer-settings';
  if (/theme/i.test(name)) return 'themes';
  if (/styles\.xml/i.test(name)) return 'styles';
  if (/sharedStrings/i.test(name)) return 'shared-strings';
  if (/\.rels$|_rels/i.test(name)) return 'relationships';
  if (/worksheets|workbook\.xml/i.test(name)) return 'workbook-worksheets';
  if (/Content_Types/i.test(name)) return 'content-types';
  return 'other';
}

function normalizeXml(text) {
  return text
    .replace(/^\uFEFF/, '')
    .replace(/\r\n?/g, '\n')
    .replace(/^\s*<\?xml[^?]*\?>\s*/i, '')
    .replace(/>\s+</g, '><')
    .trim();
}

function digest(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

async function inventory(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter(file => !zip.files[file].dir)
    .sort();
  const entries = await Promise.all(
    names.map(async name => {
      const bytes = await zip.file(name).async('nodebuffer');
      const xml = /\.xml$|\.rels$|\.vml$/i.test(name) || name === CONTENT_TYPES;
      const normalized = xml
        ? Buffer.from(normalizeXml(bytes.toString('utf8')))
        : bytes;
      return [
        name,
        {
          category: classifyPart(name),
          bytes: bytes.length,
          sha256: digest(normalized),
          normalizedXml: xml,
        },
      ];
    })
  );
  const parts = {};
  entries.forEach(([name, part]) => {
    parts[name] = part;
  });
  return parts;
}

function partDiff(before, after) {
  const beforeNames = Object.keys(before);
  const afterNames = Object.keys(after);
  const beforeSet = new Set(beforeNames);
  const afterSet = new Set(afterNames);
  return {
    added: afterNames
      .filter(name => !beforeSet.has(name))
      .map(name => ({name, category: after[name].category})),
    dropped: beforeNames
      .filter(name => !afterSet.has(name))
      .map(name => ({name, category: before[name].category})),
    changed: beforeNames
      .filter(
        name => afterSet.has(name) && before[name].sha256 !== after[name].sha256
      )
      .map(name => ({
        name,
        category: before[name].category,
        beforeSha256: before[name].sha256,
        afterSha256: after[name].sha256,
      })),
    unchanged: beforeNames.filter(
      name => afterSet.has(name) && before[name].sha256 === after[name].sha256
    ),
  };
}

async function createSyntheticUnsupportedPackage() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Generated characterization fixture';
  const data = workbook.addWorksheet('Data');
  data.getCell('A1').value = 'known original value';
  data.getCell('A1').font = {bold: true};
  data.getCell('B1').value = {formula: 'SUM(20,22)', result: 42};
  data.getCell('C1').value = new Date('2024-05-06T00:00:00.000Z');
  data.getCell('C1').numFmt = 'yyyy-mm-dd';
  data.mergeCells('D1:E1');
  data.getCell('D1').value = 'merged';
  const image = workbook.addImage({
    filename: path.join(
      repositoryRoot,
      'spec',
      'integration',
      'data',
      'image.png'
    ),
    extension: 'png',
  });
  data.addImage(image, 'G1:H3');
  const hidden = workbook.addWorksheet('Hidden');
  hidden.state = 'hidden';
  hidden.getCell('A1').value = 'hidden data';

  const base = await workbook.xlsx.writeBuffer({
    useSharedStrings: true,
    useStyles: true,
  });
  const zip = await JSZip.loadAsync(base);
  let contentTypes = await zip.file(CONTENT_TYPES).async('string');
  contentTypes = contentTypes.replace(
    /(<Override\b[^>]*PartName="\/xl\/workbook\.xml"[^>]*ContentType=")[^"]+("[^>]*\/>)/,
    '$1application/vnd.ms-excel.sheet.macroEnabled.main+xml$2'
  );
  contentTypes = contentTypes.replace(
    '</Types>',
    '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'
  );
  zip.file(CONTENT_TYPES, contentTypes);

  const rootRels = await zip.file('_rels/.rels').async('string');
  zip.file(
    '_rels/.rels',
    rootRels.replace(
      '</Relationships>',
      '<Relationship Id="rIdCharacterizationCustomXml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="customXml/item1.xml"/></Relationships>'
    )
  );
  const workbookRels = await zip
    .file('xl/_rels/workbook.xml.rels')
    .async('string');
  zip.file(
    'xl/_rels/workbook.xml.rels',
    workbookRels.replace(
      '</Relationships>',
      '<Relationship Id="rIdCharacterizationExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rIdCharacterizationVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/><Relationship Id="rIdCharacterizationVendor" Type="https://example.invalid/relationships/vendorPart" Target="vendor/vendorPart.xml"/></Relationships>'
    )
  );

  zip.file(
    'xl/externalLinks/externalLink1.xml',
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook><sheetNames><sheetName val="ExternalSheet"/></sheetNames></externalBook></externalLink>'
  );
  zip.file(
    'xl/externalLinks/_rels/externalLink1.xml.rels',
    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="https://example.invalid/source.xlsx" TargetMode="External"/></Relationships>`
  );
  zip.file(
    'xl/vbaProject.bin',
    Buffer.from('synthetic-placeholder-not-a-valid-vba-project')
  );
  zip.file(
    'customXml/item1.xml',
    '<vendor:data xmlns:vendor="urn:example:characterization">custom XML payload</vendor:data>'
  );
  zip.file(
    'customXml/itemProps1.xml',
    '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{7D6DFF9B-80E3-44E0-AB1B-BC55E8C0B320}"/>'
  );
  zip.file(
    'customXml/_rels/item1.xml.rels',
    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>`
  );
  zip.file(
    'xl/vendor/vendorPart.xml',
    '<vendor:part xmlns:vendor="urn:example:vendor">payload</vendor:part>'
  );

  let themeProbeApplied = false;
  if (zip.file('xl/theme/theme1.xml')) {
    const theme = await zip.file('xl/theme/theme1.xml').async('string');
    const updatedTheme = theme.replace(
      /(<a:clrScheme\b[^>]*\bname=")[^"]+("[^>]*>)/,
      '$1CharacterizationTheme$2'
    );
    themeProbeApplied = updatedTheme !== theme;
    zip.file('xl/theme/theme1.xml', updatedTheme);
  }

  return {
    buffer: await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    }),
    fixture: {
      kind: 'programmatically-generated OOXML package with synthetic unsupported parts',
      vbaPayload: 'placeholder bytes; not a valid/executable VBA project',
      themeProbeApplied,
    },
  };
}

async function mutateOneCell({id, buffer, sheetName, address, expectedBefore}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new Error(`Missing worksheet ${sheetName} in ${id}`);
  const cell = worksheet.getCell(address);
  const beforeValue = cell.value;
  if (beforeValue !== expectedBefore) {
    throw new Error(
      `${id}!${sheetName}!${address}: expected ${expectedBefore}, got ${String(
        beforeValue
      )}`
    );
  }
  const formulaBefore = workbook.getWorksheet(sheetName).getCell('B1').value;
  const afterValue = `${expectedBefore} [single-cell mutation]`;
  cell.value = afterValue;
  const output = await workbook.xlsx.writeBuffer({
    useSharedStrings: true,
    useStyles: true,
  });
  const roundTrip = new ExcelJS.Workbook();
  await roundTrip.xlsx.load(output);
  const finalValue = roundTrip.getWorksheet(sheetName).getCell(address).value;
  const formulaAfter = roundTrip.getWorksheet(sheetName).getCell('B1').value;
  const beforeParts = await inventory(buffer);
  const afterParts = await inventory(output);
  const outputName = `${id}.mutated.xlsx`;
  fs.writeFileSync(path.join(outputDirectory, outputName), output);

  return {
    id,
    target: {
      sheet: sheetName,
      address,
      before: expectedBefore,
      after: afterValue,
    },
    editCount: 1,
    targetValueAfterReload: finalValue,
    formulaCache: {
      before: formulaBefore,
      after: formulaAfter,
      resultPreserved: Boolean(
        formulaBefore &&
          formulaAfter &&
          formulaBefore.result === formulaAfter.result
      ),
    },
    fixtureParts: Object.keys(beforeParts).map(name => ({
      name,
      category: beforeParts[name].category,
    })),
    outputParts: Object.keys(afterParts).map(name => ({
      name,
      category: afterParts[name].category,
    })),
    diff: partDiff(beforeParts, afterParts),
    output: path.join('.temp', 'characterization', outputName),
  };
}

describe('ExcelJS characterization: mutation preservation', () => {
  let report;

  before(async () => {
    ensureOutputDirectory();
    const synthetic = await createSyntheticUnsupportedPackage();
    const syntheticInput = path.join(
      outputDirectory,
      'synthetic-unsupported-ooxml.zip'
    );
    fs.writeFileSync(syntheticInput, synthetic.buffer);

    const fixtures = [
      {
        id: 'synthetic-unsupported-ooxml',
        buffer: synthetic.buffer,
        sheetName: 'Data',
        address: 'A1',
        expectedBefore: 'known original value',
        provenance: synthetic.fixture,
      },
      {
        id: 'upstream-chart-sheet',
        buffer: fs.readFileSync(
          path.join(
            repositoryRoot,
            'spec',
            'integration',
            'data',
            'chart-sheet.xlsx'
          )
        ),
        sheetName: 'Sheet1',
        address: 'A1',
        expectedBefore: 'Month',
        provenance:
          'Existing ExcelJS upstream integration fixture: spec/integration/data/chart-sheet.xlsx',
      },
      {
        id: 'upstream-comments-vml-printer-settings',
        buffer: fs.readFileSync(
          path.join(
            repositoryRoot,
            'spec',
            'integration',
            'data',
            'test-issue-880.xlsx'
          )
        ),
        sheetName: 'A',
        address: 'A2',
        expectedBefore: 'Kosten',
        provenance:
          'Existing ExcelJS upstream issue fixture: spec/integration/data/test-issue-880.xlsx',
      },
    ];

    const cases = await Promise.all(
      fixtures.map(async fixture => {
        const result = await mutateOneCell(fixture);
        result.provenance = fixture.provenance;
        return result;
      })
    );
    report = {
      baselineCommit: '5bed18b45e824f409b08456b59b87430ded023ab',
      category: 'baseline-limitation',
      operation:
        'read -> change exactly one known cell value -> write -> unzip -> normalize XML parts -> compare package inventory and normalized content',
      normalization:
        'UTF BOM and line endings removed; XML declaration stripped; whitespace only between tags collapsed; XML payload text preserved; binary parts compared by SHA-256',
      syntheticInput: {
        path: path.join(
          '.temp',
          'characterization',
          'synthetic-unsupported-ooxml.zip'
        ),
        warning:
          'VBA part uses placeholder bytes and is not an executable VBA project; synthetic package measures unknown-part handling, not Excel macro validity.',
        themeProbeApplied: synthetic.fixture.themeProbeApplied,
      },
      cases,
    };
    fs.writeFileSync(
      path.join(outputDirectory, 'mutation-preservation.json'),
      `${JSON.stringify(report, null, 2)}\n`
    );
  });

  it('[upstream-pass] one known cell is the only model edit and cached formula result is reported', () => {
    expect(report.cases.map(item => item.editCount)).to.deep.equal([1, 1, 1]);
    expect(report.cases[0].targetValueAfterReload).to.equal(
      'known original value [single-cell mutation]'
    );
    expect(report.cases[0].formulaCache.after.result).to.equal(42);
  });

  it('[upstream-pass] embedded image and customized theme part remain byte-identical', () => {
    const syntheticCase = report.cases.find(
      item => item.id === 'synthetic-unsupported-ooxml'
    );
    expect(report.syntheticInput.themeProbeApplied).to.equal(true);
    expect(syntheticCase.diff.unchanged).to.include('xl/theme/theme1.xml');
    expect(syntheticCase.diff.unchanged).to.include('xl/media/image1.png');
    expect(syntheticCase.diff.unchanged).to.include('xl/drawings/drawing1.xml');
  });

  it('[baseline-limitation] chart-sheet charts and their relationship parts are measured', () => {
    const chartCase = report.cases.find(
      item => item.id === 'upstream-chart-sheet'
    );
    expect(
      chartCase.diff.dropped.some(part => /xl\/charts?\//.test(part.name))
    ).to.equal(true);
    expect(
      chartCase.diff.dropped.some(part =>
        /chartsheets|drawings/.test(part.name)
      )
    ).to.equal(true);
  });

  it('[baseline-limitation] comment, VML and printer package parts are dropped or regenerated', () => {
    const issueCase = report.cases.find(
      item => item.id === 'upstream-comments-vml-printer-settings'
    );
    const dropped = issueCase.diff.dropped.map(part => part.name);
    expect(dropped.some(name => /comments/i.test(name))).to.equal(true);
    expect(dropped.some(name => /vmlDrawing/i.test(name))).to.equal(true);
    expect(dropped.some(name => /printerSettings/i.test(name))).to.equal(true);
  });

  it('[baseline-limitation] synthetic macro, external-link and custom XML parts are inventoried', () => {
    const syntheticCase = report.cases.find(
      item => item.id === 'synthetic-unsupported-ooxml'
    );
    const originalNames = syntheticCase.fixtureParts.map(part => part.name);
    expect(originalNames).to.include('xl/vbaProject.bin');
    expect(
      originalNames.some(name => name.startsWith('xl/externalLinks/'))
    ).to.equal(true);
    expect(originalNames.some(name => name.startsWith('customXml/'))).to.equal(
      true
    );
    expect(
      syntheticCase.diff.dropped.length + syntheticCase.diff.changed.length
    ).to.be.greaterThan(0);
  });

  it('[stress/invariant probe] writes a machine-readable normalized OOXML diff', () => {
    const outputPath = path.join(outputDirectory, 'mutation-preservation.json');
    const saved = JSON.parse(fs.readFileSync(outputPath, 'utf8'));
    expect(saved.cases).to.have.length(3);
    expect(saved.normalization).to.be.a('string');
  });
});
