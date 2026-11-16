const XmlStream = require('../../../utils/xml-stream');
const BaseXform = require('../base-xform');
const colCache = require('../../../utils/col-cache');

class PivotTableXform extends BaseXform {
  constructor() {
    super();

    this.map = {};
  }

  prepare(model) {
    // TK
  }

  get tag() {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
    return 'pivotTableDefinition';
  }

  render(xmlStream, model) {
    // rows, columns and values[].field are indices into `cacheFields`.
    const {rows, columns, filters, values, cacheFields, cacheId} = model;

    const rowItems = renderRowItems(model);
    const multiValue = values.length > 1;
    // Compact layout: one row label column, then one column per data field.
    // With column fields, Excel lays the body out again on refresh.
    const width = columns.length ? 2 : 1 + values.length;
    const height = (columns.length ? 2 : 1) + rowItems.length;
    // Filters sit at the top, one per row, then one blank row, then the pivot.
    const top = filters.length ? filters.length + 2 : 3;
    const ref = `A${top}:${colCache.n2l(width)}${top - 1 + height}`;

    const attributes = {
      ...PivotTableXform.PIVOT_TABLE_ATTRIBUTES,
      'xr:uid': '{267EE50F-B116-784D-8DC2-BA77DE3F4F4A}',
      name: model.name,
      cacheId,
      applyNumberFormats: '0',
      applyBorderFormats: '0',
      applyFontFormats: '0',
      applyPatternFormats: '0',
      applyAlignmentFormats: '0',
      applyWidthHeightFormats: '1',
      dataCaption: 'Values',
      updatedVersion: '8',
      minRefreshableVersion: '3',
      useAutoFormatting: '1',
      itemPrintTitles: '1',
      createdVersion: '8',
      indent: '0',
      outline: '1',
      outlineData: '1',
      multipleFieldFilters: '0',
    };
    if (model.missingCaption !== undefined) attributes.missingCaption = model.missingCaption;
    if (model.rowHeaderCaption !== undefined) attributes.rowHeaderCaption = model.rowHeaderCaption;

    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, attributes);

    const location = {
      ref,
      firstHeaderRow: multiValue ? '0' : '1',
      firstDataRow: columns.length ? '2' : '1',
      firstDataCol: '1',
    };
    if (filters.length) {
      location.rowPageCount = filters.length;
      location.colPageCount = '1';
    }
    xmlStream.leafNode('location', location);

    xmlStream.writeXml(
      `<pivotFields count="${cacheFields.length}">${cacheFields
        .map((cacheField, index) => renderPivotField(model, cacheField, index))
        .join('')}</pivotFields>`
    );

    xmlStream.writeXml(
      `<rowFields count="${rows.length}">${rows
        .map(x => `<field x="${x}"/>`)
        .join('')}</rowFields>`
    );
    xmlStream.writeXml(`<rowItems count="${rowItems.length}">${rowItems.join('')}</rowItems>`);

    // Several data fields sit on the column axis as the "Values" field (x = -2).
    const colFields = multiValue ? [-2] : columns;
    if (colFields.length) {
      xmlStream.writeXml(
        `<colFields count="${colFields.length}">${colFields
          .map(x => `<field x="${x}"/>`)
          .join('')}</colFields>`
      );
    }
    let colItems = ['<i/>'];
    if (multiValue)
      colItems = values.map((v, i) => (i ? `<i i="${i}"><x v="${i}"/></i>` : '<i><x/></i>'));
    // ponytail: column-field items are left to refreshOnLoad; compute them when a file must render without refresh.
    if (columns.length) colItems = ['<i t="grand"><x/></i>'];
    xmlStream.writeXml(`<colItems count="${colItems.length}">${colItems.join('')}</colItems>`);

    // No `item` attribute: the filter shows "(All)".
    if (filters.length) {
      xmlStream.writeXml(
        `<pageFields count="${filters.length}">${filters
          .map(fld => `<pageField fld="${fld}" hier="-1"/>`)
          .join('')}</pageFields>`
      );
    }

    xmlStream.openNode('dataFields', {count: values.length});
    values.forEach(v => {
      const dataField = {
        name: v.caption,
        fld: v.field,
        baseField: '0',
        baseItem: '0',
      };
      if (v.metric !== 'sum') dataField.subtotal = v.metric;
      if (v.numFmtId !== undefined) dataField.numFmtId = v.numFmtId;
      xmlStream.leafNode('dataField', dataField);
    });
    xmlStream.closeNode();

    xmlStream.leafNode('pivotTableStyleInfo', {
      name: 'PivotStyleLight16',
      showRowHeaders: '1',
      showColHeaders: '1',
      showRowStripes: '0',
      showColStripes: '0',
      showLastColumn: '1',
    });

    xmlStream.closeNode();
  }

  parseOpen(node) {
    // TK
  }

  parseText(text) {
    // TK
  }

  parseClose(name) {
    // TK
  }

  reconcile(model, options) {
    // TK
  }
}

// Helpers

function renderPivotField(model, cacheField, index) {
  let axis = null;
  if (model.rows.includes(index)) axis = 'axisRow';
  else if (model.columns.includes(index)) axis = 'axisCol';
  else if (model.filters.includes(index)) axis = 'axisPage';
  const dataField = model.values.some(v => v.field === index) ? ' dataField="1"' : '';
  if (!axis) return `<pivotField${dataField} showAll="0"/>`;

  // Outer row fields start collapsed when asked (sd = show details).
  const isOuterRow = axis === 'axisRow' && model.rows.indexOf(index) < model.rows.length - 1;
  const sd = model.collapsed && isOuterRow ? ' sd="0"' : '';
  const items = cacheField.sharedItems.map((item, i) => `<item x="${i}"${sd}/>`);
  items.push('<item t="default"/>');

  const sort = model.sort.find(s => s.field === index);
  const sortType = sort ? ` sortType="${sort.order}"` : '';
  // Sort by a data field: the "Values" field (4294967294 = -2) at that data field.
  const scope =
    sort && sort.by !== undefined
      ? '<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0">' +
        '<references count="1"><reference field="4294967294" count="1" selected="0">' +
        `<x v="${sort.by}"/></reference></references></pivotArea></autoSortScope>`
      : '';
  return `<pivotField axis="${axis}"${dataField} showAll="0"${sortType}><items count="${
    items.length
  }">${items.join('')}</items>${scope}</pivotField>`;
}

// One data field's value over a group of records (as Excel computes it).
function aggregate(records, value) {
  const xs = records.map(r => r[value.field]).filter(v => v !== undefined);
  const nums = xs.filter(v => typeof v === 'number');
  switch (value.metric) {
    case 'count':
      return xs.length;
    case 'countNums':
      return nums.length;
    case 'average':
      return nums.length ? nums.reduce((t, v) => t + v, 0) / nums.length : 0;
    case 'max':
      return nums.length ? nums.reduce((m, v) => (v > m ? v : m)) : 0;
    case 'min':
      return nums.length ? nums.reduce((m, v) => (v < m ? v : m)) : 0;
    default:
      return nums.reduce((t, v) => t + v, 0);
  }
}

// Compact-form row lines: each distinct path through the row fields, skipping
// children of collapsed items; then grand total. Items follow the field's
// sort (by label, or by a data field), else shared-item order.
function renderRowItems(model) {
  const {rows, records, cacheFields, collapsed, sort, values} = model;
  const tree = {children: new Map(), records};
  records.forEach(record => {
    let node = tree;
    rows.forEach(field => {
      const x = cacheFields[field].sharedItems.indexOf(record[field]);
      if (!node.children.has(x)) node.children.set(x, {children: new Map(), records: []});
      node = node.children.get(x);
      node.records.push(record);
    });
  });

  const order = (node, field) => {
    const keys = [...node.children.keys()].sort((a, b) => a - b);
    const s = sort.find(f => f.field === field);
    if (!s) return keys;
    const sign = s.order === 'descending' ? -1 : 1;
    const key =
      s.by === undefined
        ? x => x // shared items are already in label order
        : x => aggregate(node.children.get(x).records, values[s.by]);
    // Stable: ties keep label order.
    return keys
      .map(x => [x, key(x)])
      .sort((p, q) => sign * (p[1] - q[1]))
      .map(([x]) => x);
  };

  const lines = [];
  const walk = (node, depth) => {
    order(node, rows[depth]).forEach(x => {
      lines.push(`<i${depth ? ` r="${depth}"` : ''}>${x ? `<x v="${x}"/>` : '<x/>'}</i>`);
      if (!collapsed && depth + 1 < rows.length) walk(node.children.get(x), depth + 1);
    });
  };
  walk(tree, 0);
  lines.push('<i t="grand"><x/></i>');
  return lines;
}

PivotTableXform.PIVOT_TABLE_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'mc:Ignorable': 'xr',
  'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
};

module.exports = PivotTableXform;
