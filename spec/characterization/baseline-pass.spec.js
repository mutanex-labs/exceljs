'use strict';

const path = require('path');
const {
  ExcelJS,
  JSZip,
  chunkReadable,
  readStreaming,
  repositoryRoot,
  writeStreamingSheets,
  writeStreamingWorkbook,
} = require('./helpers');

function createFeatureWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ExcelJS upstream characterization';
  workbook.created = new Date('2020-01-02T03:04:05.000Z');
  workbook.modified = new Date('2020-01-02T03:04:05.000Z');
  const sheet = workbook.addWorksheet('Primary');
  sheet.getCell('A1').value = 'Shared heading';
  sheet.getCell('A1').font = {bold: true, color: {argb: 'FF123456'}};
  sheet.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {argb: 'FFE0E0E0'},
  };
  sheet.mergeCells('A3:B3');
  sheet.getCell('A3').value = 'Merged label';
  sheet.getCell('A4').value = 1;
  sheet.getCell('A5').value = 2;
  sheet.getCell('B1').value = {formula: 'SUM(A4:A5)', result: 3};
  sheet.getCell('C1').value = {
    richText: [
      {font: {bold: true}, text: 'rich '},
      {font: {italic: true}, text: 'text'},
    ],
  };
  sheet.getCell('D1').value = {
    text: 'External site',
    hyperlink: 'https://example.invalid/workbook',
  };
  sheet.getCell('E1').value = new Date('2021-02-03T04:05:06.000Z');
  sheet.getCell('E1').numFmt = 'yyyy-mm-dd hh:mm:ss';
  sheet.getCell('F1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"small,large"'],
  };
  sheet.addConditionalFormatting({
    ref: 'A4:A5',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [1],
        style: {font: {bold: true}},
      },
    ],
  });
  sheet.getCell('A1').addName('PrimaryHeading');

  const hidden = workbook.addWorksheet('Hidden');
  hidden.state = 'hidden';
  hidden.getCell('A1').value = 'hidden row';
  const veryHidden = workbook.addWorksheet('Very Hidden');
  veryHidden.state = 'veryHidden';
  veryHidden.getCell('A1').value = 'very hidden row';

  return workbook;
}

describe('ExcelJS characterization: baseline pass', () => {
  it('[upstream-pass] ordinary XLSX round-trip covers common workbook features', async () => {
    const workbook = createFeatureWorkbook();
    const imagePath = path.join(
      repositoryRoot,
      'spec',
      'integration',
      'data',
      'image.png'
    );
    const imageId = workbook.addImage({filename: imagePath, extension: 'png'});
    workbook.getWorksheet('Primary').addImage(imageId, 'H2:I4');

    const buffer = await workbook.xlsx.writeBuffer({
      useSharedStrings: true,
      useStyles: true,
    });
    const zip = await JSZip.loadAsync(buffer);
    const output = new ExcelJS.Workbook();
    await output.xlsx.load(buffer);

    expect(Object.keys(zip.files)).to.include('xl/sharedStrings.xml');
    expect(output.worksheets.map(sheet => sheet.name)).to.deep.equal([
      'Primary',
      'Hidden',
      'Very Hidden',
    ]);
    const primary = output.getWorksheet('Primary');
    expect(primary.getCell('A1').value).to.equal('Shared heading');
    expect(primary.getCell('A1').font.bold).to.equal(true);
    expect(primary.getCell('A1').fill.fgColor.argb).to.equal('FFE0E0E0');
    expect(primary.getCell('B3').master.address).to.equal('A3');
    expect(primary.getCell('B1').value).to.deep.equal({
      formula: 'SUM(A4:A5)',
      result: 3,
    });
    expect(
      primary.getCell('C1').value.richText.map(run => run.text)
    ).to.deep.equal(['rich ', 'text']);
    expect(primary.getCell('D1').value.hyperlink).to.equal(
      'https://example.invalid/workbook'
    );
    expect(primary.getCell('E1').value.toISOString()).to.equal(
      '2021-02-03T04:05:06.000Z'
    );
    expect(primary.getCell('F1').dataValidation.formulae).to.deep.equal([
      '"small,large"',
    ]);
    expect(primary.conditionalFormattings).to.have.length(1);
    expect(primary.getCell('A1').names).to.include('PrimaryHeading');
    expect(output.getWorksheet('Hidden').state).to.equal('hidden');
    expect(output.getWorksheet('Very Hidden').state).to.equal('veryHidden');
    expect(primary.getImages()).to.have.length(1);
    expect(
      output.getImage(primary.getImages()[0].imageId).buffer
    ).to.be.instanceOf(Buffer);
  });

  it('[upstream-pass] 1900 and 1904 date systems round-trip dates', async () => {
    const results = await Promise.all(
      [false, true].map(async date1904 => {
        const workbook = new ExcelJS.Workbook();
        workbook.properties.date1904 = date1904;
        const worksheet = workbook.addWorksheet('Dates');
        worksheet.getCell('A1').value = new Date('1904-01-01T00:00:00.000Z');
        const buffer = await workbook.xlsx.writeBuffer({useStyles: true});
        const output = new ExcelJS.Workbook();
        await output.xlsx.load(buffer);

        return {date1904, output};
      })
    );
    for (const {date1904, output} of results) {
      expect(output.properties.date1904).to.equal(date1904);
      expect(
        output.getWorksheet('Dates').getCell('A1').value.toISOString()
      ).to.equal('1904-01-01T00:00:00.000Z');
    }
  });

  it('[upstream-pass] WorkbookReader emits multiple worksheets and shared-string rows', async () => {
    const buffer = await writeStreamingSheets([
      {name: 'First', rows: [['Shared heading'], ['first row']]},
      {name: 'Second', rows: [['another value']]},
    ]);
    const result = await readStreaming(chunkReadable(buffer, 1));

    expect(result.status).to.equal('end');
    expect(result.worksheets.map(sheet => sheet.name)).to.deep.equal([
      'First',
      'Second',
    ]);
    expect(
      result.rows.some(row =>
        row.values.some(value => value.value === 'Shared heading')
      )
    ).to.equal(true);
  });

  it('[upstream-pass] WorkbookWriter output can be read back', async () => {
    const buffer = await writeStreamingWorkbook([
      [1, 'first'],
      [2, 'second'],
    ]);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    expect(
      workbook
        .getWorksheet('Data')
        .getSheetValues()
        .slice(1)
        .map(row => row.slice(1))
    ).to.deep.equal([
      [1, 'first'],
      [2, 'second'],
    ]);
  });

  it('[baseline-limitation] CSV defaults parse dates back as strings', async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('CSV');
    worksheet.addRow(['heading', 'count', 'date']);
    worksheet.addRow(['alpha, beta', 7, new Date('2024-03-04T00:00:00.000Z')]);
    const csv = await workbook.csv.writeBuffer({
      sheetName: 'CSV',
      dateUTC: true,
    });

    const output = new ExcelJS.Workbook();
    const parsed = await output.csv.read(
      chunkReadable(Buffer.from(csv), Buffer.byteLength(csv)),
      {
        sheetName: 'Parsed',
      }
    );
    expect(parsed.getRow(1).values.slice(1)).to.deep.equal([
      'heading',
      'count',
      'date',
    ]);
    expect(parsed.getRow(2).values.slice(1)).to.deep.equal([
      'alpha, beta',
      7,
      '2024-03-04T00:00:00Z',
    ]);
  });

  it('[stress/invariant probe] streaming reader handles the upstream large workbook', async function() {
    this.timeout(30000);
    const filename = path.join(
      repositoryRoot,
      'spec',
      'integration',
      'data',
      'huge.xlsx'
    );
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(filename, {
      worksheets: 'emit',
      sharedStrings: 'cache',
      styles: 'ignore',
      hyperlinks: 'ignore',
    });
    let rowCount = 0;
    let firstRow;
    let lastRow;
    for await (const worksheet of reader) {
      if (worksheet.name !== '100k') continue;
      for await (const row of worksheet) {
        rowCount += 1;
        if (rowCount === 1) firstRow = row.values.slice(1);
        lastRow = row.values.slice(1);
      }
    }
    expect(rowCount).to.equal(100001);
    expect(firstRow).to.have.length.greaterThan(0);
    expect(lastRow).to.have.length.greaterThan(0);
  });

  it('[stress/invariant probe] WorkbookWriter streams and reads back 100,000 rows', async function() {
    this.timeout(60000);
    const buffer = await writeStreamingWorkbook(100000, {
      useSharedStrings: false,
      useStyles: false,
    });
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(
      chunkReadable(buffer, buffer.length),
      {
        worksheets: 'emit',
        sharedStrings: 'ignore',
        styles: 'ignore',
        hyperlinks: 'ignore',
      }
    );
    let rowCount = 0;
    let first;
    let last;
    for await (const worksheet of reader) {
      for await (const row of worksheet) {
        rowCount += 1;
        if (rowCount === 1) first = row.values.slice(1);
        last = row.values.slice(1);
      }
    }
    expect(rowCount).to.equal(100000);
    expect(first).to.deep.equal([1, 'row-1']);
    expect(last).to.deep.equal([100000, 'row-100000']);
  });
});
