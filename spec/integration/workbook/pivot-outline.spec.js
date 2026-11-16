const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const DATA = [
  ['Item', 'Description', 'Job', 'Qty'],
  ['A-100', 'Bolt & nut', 'J-1', 4],
  ['A-100', 'Bolt & nut', 'J-2', 6],
  ['B-200', 'Washer', null, null],
  ['C-300', 'Pipe', 'J-2', 12.5],
];

function jobsWorkbook(options = {}) {
  const workbook = new ExcelJS.Workbook();
  const data = workbook.addWorksheet('Data');
  data.addRows(DATA);
  workbook.addWorksheet('Pivot').addPivotTable({
    sourceSheet: data,
    rows: ['Item', 'Job'],
    values: [
      {name: 'Job', metric: 'count', caption: 'Jobs'},
      {name: 'Qty', metric: 'sum'},
    ],
    missingCaption: '0',
    ...options,
  });
  return workbook;
}

async function parts(workbook) {
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  return name => (zip.file(name) ? zip.file(name).async('string') : undefined);
}

describe('Workbook', () => {
  describe('Pivot table with row hierarchy and several values', () => {
    it('writes row fields, data fields and compact row items', async () => {
      const read = await parts(jobsWorkbook());
      const xml = await read('xl/pivotTables/pivotTable1.xml');

      expect(xml).to.contain('<rowFields count="2"><field x="0"/><field x="2"/></rowFields>');
      // A-100 > J-1, J-2 ; B-200 > (blank) ; C-300 > J-2 ; grand
      expect(xml).to.contain(
        '<rowItems count="8"><i><x/></i><i r="1"><x/></i><i r="1"><x v="1"/></i>' +
          '<i><x v="1"/></i><i r="1"><x v="2"/></i>' +
          '<i><x v="2"/></i><i r="1"><x v="1"/></i><i t="grand"><x/></i></rowItems>'
      );
      expect(xml).to.contain('<colFields count="1"><field x="-2"/></colFields>');
      expect(xml).to.contain('<colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>');
      expect(xml).to.contain(
        '<dataField name="Jobs" fld="2" baseField="0" baseItem="0" subtotal="count"/>'
      );
      expect(xml).to.contain('<dataField name="Sum of Qty" fld="3" baseField="0" baseItem="0"/>');
      expect(xml).to.contain(
        '<location ref="A3:C11" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>'
      );
      expect(xml).to.contain('missingCaption="0"');
      // Job is both a row field and a data field.
      expect(xml).to.contain('<pivotField axis="axisRow" dataField="1" showAll="0">');
      expect(xml).not.to.contain('sd="0"');
    });

    it('collapses the outer row items when asked', async () => {
      const read = await parts(jobsWorkbook({collapsed: true}));
      const xml = await read('xl/pivotTables/pivotTable1.xml');

      expect(xml).to.contain(
        '<item x="0" sd="0"/><item x="1" sd="0"/><item x="2" sd="0"/><item t="default"/>'
      );
      expect(xml).to.contain(
        '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems>'
      );
    });

    it('writes a refreshable cache with blanks and escaped text', async () => {
      const read = await parts(jobsWorkbook());
      const definition = await read('xl/pivotCache/pivotCacheDefinition1.xml');
      const records = await read('xl/pivotCache/pivotCacheRecords1.xml');

      expect(definition).to.contain('refreshOnLoad="1"');
      expect(definition).to.contain('recordCount="4"');
      expect(definition).to.contain('<worksheetSource ref="A1:D5" sheet="Data"/>');
      expect(definition).to.contain(
        '<sharedItems containsBlank="1" count="3"><s v="J-1"/><s v="J-2"/><m/></sharedItems>'
      );
      expect(definition).to.contain(
        '<sharedItems containsString="0" containsNumber="1" minValue="4" maxValue="12.5" containsBlank="1"/>'
      );

      expect(records).to.contain('count="4"');
      expect(records).to.contain('<r><x v="0"/><s v="Bolt &amp; nut"/><x v="0"/><n v="4"/></r>');
      expect(records).to.contain('<r><x v="1"/><s v="Washer"/><x v="2"/><m/></r>');
    });

    it('links the parts through rels, workbook and content types', async () => {
      const read = await parts(jobsWorkbook());

      const workbookXml = await read('xl/workbook.xml');
      const workbookRels = await read('xl/_rels/workbook.xml.rels');
      const rId = workbookXml.match(/<pivotCache cacheId="10" r:id="(rId\d+)"\/>/)[1];
      expect(workbookRels).to.match(
        new RegExp(
          `Id="${rId}" Type="[^"]+/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"`
        )
      );
      expect(await read('xl/pivotTables/pivotTable1.xml')).to.contain('cacheId="10"');

      expect(await read('xl/worksheets/_rels/sheet2.xml.rels')).to.contain(
        'Target="../pivotTables/pivotTable1.xml"'
      );
      expect(await read('xl/worksheets/_rels/sheet1.xml.rels')).to.be.undefined();
      expect(await read('xl/pivotTables/_rels/pivotTable1.xml.rels')).to.contain(
        'Target="../pivotCache/pivotCacheDefinition1.xml"'
      );
      expect(await read('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')).to.contain(
        'Target="pivotCacheRecords1.xml"'
      );

      const types = await read('[Content_Types].xml');
      [
        ['/xl/pivotCache/pivotCacheDefinition1.xml', 'pivotCacheDefinition'],
        ['/xl/pivotCache/pivotCacheRecords1.xml', 'pivotCacheRecords'],
        ['/xl/pivotTables/pivotTable1.xml', 'pivotTable'],
      ].forEach(([part, type]) => {
        expect(types).to.contain(
          `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${type}+xml"/>`
        );
      });
    });

    it('keeps a data field caption apart from the source field names', async () => {
      const read = await parts(jobsWorkbook({values: [{name: 'Qty', caption: 'Qty'}]}));
      expect(await read('xl/pivotTables/pivotTable1.xml')).to.contain(
        '<dataField name="Qty " fld="3"'
      );
    });

    it('rejects an unknown metric and several values with column fields', () => {
      expect(() => jobsWorkbook({values: [{name: 'Qty', metric: 'median'}]})).to.throw(
        /Unknown metric/
      );
      expect(() => jobsWorkbook({columns: ['Description']})).to.throw(/column fields/);
    });
  });

  describe('Row outline', () => {
    it('writes a collapsed summary row above its hidden detail rows', async () => {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Grouped');
      sheet.properties.outlineProperties = {summaryBelow: false};
      sheet.addRow(['Item', 'Jobs']);
      const summary = sheet.addRow(['A-100', 2]);
      summary.collapsed = true;
      ['J-1', 'J-2'].forEach(job => {
        const detail = sheet.addRow([job]);
        detail.outlineLevel = 1;
        detail.hidden = true;
        detail.collapsed = false;
      });

      const read = await parts(workbook);
      const xml = await read('xl/worksheets/sheet1.xml');

      expect(xml).to.contain('<outlinePr summaryBelow="0"/>');
      expect(xml).to.contain('outlineLevelRow="1"');
      expect(xml).to.match(/<row r="2"[^>]* collapsed="1"/);
      expect(xml).to.match(/<row r="3" hidden="1"[^>]* outlineLevel="1">/);
      expect(xml).not.to.match(/<row r="3"[^>]*collapsed/);
    });
  });
});
