/* eslint-disable */
// Writes a sample with a pivot (rows Item > Job; Job count, Qty sum) and an outlined sheet.
// Usage: node test/test-pivot-jobs.js out.xlsx
const Excel = require('../lib/exceljs.nodejs.js');

const data = [
  ['Item', 'Description', 'Job', 'Reserved Qty'],
  ['A-100', 'Bolt & nut', 'J-1', 4],
  ['A-100', 'Bolt & nut', 'J-2', 6],
  ['A-100', 'Bolt & nut', 'J-3', 1],
  ['B-200', 'Washer <M8>', null, null],
  ['C-300', 'Pipe "2in"', 'J-2', 12.5],
];

const wb = new Excel.Workbook();
const ws = wb.addWorksheet('Data');
ws.addRows(data);

wb.addWorksheet('Pivot').addPivotTable({
  sourceSheet: ws,
  rows: ['Item', 'Job'],
  values: [
    {name: 'Job', metric: 'count', caption: 'Jobs'},
    {name: 'Reserved Qty', metric: 'sum'},
  ],
  collapsed: true,
  missingCaption: '0',
  rowHeaderCaption: 'Item / Job',
});

const grouped = wb.addWorksheet('Grouped');
grouped.properties.outlineProperties = {summaryBelow: false};
grouped.addRow(['Item', 'Jobs', 'Reserved Qty']);
const add = (values, level, hidden, collapsed) => {
  const row = grouped.addRow(values);
  row.outlineLevel = level;
  row.hidden = hidden;
  row.collapsed = collapsed;
};
add(['A-100', 3, 11], 0, false, true);
add(['  J-1', null, 4], 1, true, false);
add(['  J-2', null, 6], 1, true, false);
add(['  J-3', null, 1], 1, true, false);
add(['B-200', 0, 0], 0, false, false);
add(['C-300', 1, 12.5], 0, false, false);
add(['  J-2', null, 12.5], 1, false, false);

wb.xlsx.writeFile(process.argv[2]).then(() => console.log('wrote', process.argv[2]));
