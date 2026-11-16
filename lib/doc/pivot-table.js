const {objectFromProps, toSortedArray} = require('../utils/utils');
const colCache = require('../utils/col-cache');

// Excel data-field functions (ST_DataConsolidateFunction).
const METRICS = [
  'sum',
  'count',
  'average',
  'max',
  'min',
  'product',
  'countNums',
  'stdDev',
  'stdDevp',
  'var',
  'varp',
];

const METRIC_LABEL = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
};

/**
 * Build the pivot table model from the user configuration.
 *
 * {
 *   sourceSheet: worksheet,          // header names in row 1, data below
 *   rows: ['Item', 'Job'],           // row fields, outer first
 *   columns: [],                     // optional column fields
 *   values: ['Qty'] or [{name: 'Job', metric: 'count', caption: 'Jobs'}],
 *   metric: 'sum',                   // default metric for string values
 *   collapsed: false,                // outer row items start collapsed
 *   missingCaption: '0',             // text for empty data cells
 *   rowHeaderCaption: 'Item',        // header of the compact row column
 *   filters: ['Customer'],           // page fields above the pivot, "(All)" selected
 *   sort: [{field: 'Item', by: 'Jobs', order: 'descending'}], // by = value caption or field
 *   name: 'PivotTable1',
 * }
 */
function makePivotTable(worksheet, model) {
  validate(worksheet, model);

  const {sourceSheet} = model;
  const columns = model.columns || [];
  const filters = model.filters || [];
  const valueSpecs = model.values.map(v =>
    typeof v === 'string' ? {name: v, metric: model.metric || 'sum'} : {metric: 'sum', ...v}
  );

  const names = sourceSheet.getRow(1).values.slice(1);
  const axisNames = objectFromProps([...model.rows, ...columns, ...filters], true);
  const records = readRecords(sourceSheet, names.length);
  const cacheFields = names.map((name, i) =>
    makeCacheField(
      name,
      records.map(r => r[i]),
      !!axisNames[name]
    )
  );

  const nameToIndex = names.reduce((result, name, index) => {
    result[name] = index;
    return result;
  }, {});

  const values = valueSpecs.map(v => {
    const label = METRIC_LABEL[v.metric] || v.metric;
    let caption = v.caption || `${label} of ${v.name}`;
    // Excel rejects a data field caption equal to a source field name.
    while (nameToIndex[caption] !== undefined) caption += ' ';
    return {
      field: nameToIndex[v.name],
      metric: v.metric,
      caption,
      numFmtId: v.numFmtId,
    };
  });

  const sort = (model.sort || []).map(s => ({
    field: nameToIndex[s.field],
    order: s.order || 'ascending',
    // index into `values`; undefined sorts by item label
    by: s.by === undefined ? undefined : findValue(valueSpecs, s.by),
  }));

  return {
    sourceSheet,
    sourceRef: `A1:${colCache.n2l(names.length)}${Math.max(sourceSheet.rowCount, 1)}`,
    rows: model.rows.map(name => nameToIndex[name]),
    columns: columns.map(name => nameToIndex[name]),
    filters: filters.map(name => nameToIndex[name]),
    sort,
    values,
    cacheFields,
    records,
    collapsed: !!model.collapsed,
    missingCaption: model.missingCaption,
    rowHeaderCaption: model.rowHeaderCaption,
    name: model.name || 'PivotTable1',
    // defined in <pivotTableDefinition> of xl/pivotTables/pivotTable1.xml;
    // also used in xl/workbook.xml
    cacheId: '10',
  };
}

// A sort `by` names a value by its caption or by its source field.
function findValue(valueSpecs, by) {
  const byCaption = valueSpecs.findIndex(v => v.caption === by);
  return byCaption >= 0 ? byCaption : valueSpecs.findIndex(v => v.name === by);
}

function validate(worksheet, model) {
  if (worksheet.workbook.pivotTables.length === 1) {
    throw new Error(
      'A pivot table was already added. At this time, ExcelJS supports at most one pivot table per file.'
    );
  }

  const columns = model.columns || [];
  const filters = model.filters || [];
  const valueNames = (model.values || []).map(v => (typeof v === 'string' ? v : v.name));
  const metrics = (model.values || []).map(v => (typeof v === 'string' ? model.metric : v.metric));

  for (const metric of metrics) {
    if (metric && !METRICS.includes(metric)) {
      throw new Error(`Unknown metric "${metric}". Use one of: ${METRICS.join(', ')}.`);
    }
  }

  const headerNames = model.sourceSheet.getRow(1).values.slice(1);
  const isInHeaderNames = objectFromProps(headerNames, true);
  for (const name of [...model.rows, ...columns, ...filters, ...valueNames]) {
    if (!isInHeaderNames[name]) {
      throw new Error(`The header name "${name}" was not found in ${model.sourceSheet.name}.`);
    }
  }

  if (!model.rows.length) {
    throw new Error('No pivot table rows specified.');
  }

  if (!valueNames.length) {
    throw new Error('No pivot table values specified.');
  }

  for (const name of filters) {
    if (model.rows.includes(name) || columns.includes(name)) {
      throw new Error(`"${name}" cannot be a filter and a row or column field.`);
    }
  }

  const valueSpecs = (model.values || []).map(v => (typeof v === 'string' ? {name: v} : v));
  for (const s of model.sort || []) {
    if (!model.rows.includes(s.field) && !columns.includes(s.field)) {
      throw new Error(`Sort field "${s.field}" is not a row or column field.`);
    }
    if (s.order && s.order !== 'ascending' && s.order !== 'descending') {
      throw new Error(`Sort order must be "ascending" or "descending", not "${s.order}".`);
    }
    if (s.by !== undefined && findValue(valueSpecs, s.by) < 0) {
      throw new Error(`Sort by "${s.by}" matches no value caption or field.`);
    }
  }

  if (columns.length && valueNames.length > 1) {
    throw new Error('Several values together with column fields are not supported yet.');
  }
}

// One plain value per cell; undefined means blank.
function plainValue(value) {
  if (value === null || value === undefined || value === '') return undefined;
  if (typeof value === 'number' || typeof value === 'string') return value;
  if (typeof value === 'boolean') return String(value);
  if (value.richText) return value.richText.map(t => t.text).join('');
  if (value.result !== undefined) return plainValue(value.result);
  if (value.text !== undefined) return plainValue(value.text);
  // ponytail: dates and errors become text; add <d>/<e> cache items when a pivot needs them.
  return String(value instanceof Date ? value.toISOString() : value);
}

function readRecords(sheet, fieldCount) {
  const records = [];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const record = [];
    for (let c = 1; c <= fieldCount; c++) record.push(plainValue(row.getCell(c).value));
    records.push(record);
  }
  return records;
}

function makeCacheField(name, values, isAxis) {
  const present = values.filter(v => v !== undefined);
  const numbers = present.filter(v => typeof v === 'number');
  return {
    name,
    containsBlank: present.length < values.length,
    containsNumber: numbers.length > 0,
    containsString: numbers.length < present.length,
    containsInteger: numbers.length > 0 && numbers.every(Number.isInteger),
    minValue: numbers.length ? numbers.reduce((m, v) => (v < m ? v : m)) : undefined,
    maxValue: numbers.length ? numbers.reduce((m, v) => (v > m ? v : m)) : undefined,
    // Blank sorts last, as in Excel.
    sharedItems: isAxis
      ? [...toSortedArray(new Set(present)), ...(present.length < values.length ? [undefined] : [])]
      : null,
  };
}

module.exports = {makePivotTable, METRICS};
