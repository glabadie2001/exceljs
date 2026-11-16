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
    const {rows, columns, values, cacheFields, cacheId} = model;

    const rowItems = renderRowItems(model);
    const multiValue = values.length > 1;
    // Compact layout: one row label column, then one column per data field.
    // With column fields, Excel lays the body out again on refresh.
    const width = columns.length ? 2 : 1 + values.length;
    const height = (columns.length ? 2 : 1) + rowItems.length;
    const ref = `A3:${colCache.n2l(width)}${2 + height}`;

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

    xmlStream.leafNode('location', {
      ref,
      firstHeaderRow: multiValue ? '0' : '1',
      firstDataRow: columns.length ? '2' : '1',
      firstDataCol: '1',
    });

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
  const dataField = model.values.some(v => v.field === index) ? ' dataField="1"' : '';
  if (!axis) return `<pivotField${dataField} showAll="0"/>`;

  // Outer row fields start collapsed when asked (sd = show details).
  const isOuterRow = axis === 'axisRow' && model.rows.indexOf(index) < model.rows.length - 1;
  const sd = model.collapsed && isOuterRow ? ' sd="0"' : '';
  const items = cacheField.sharedItems.map((item, i) => `<item x="${i}"${sd}/>`);
  items.push('<item t="default"/>');
  return `<pivotField axis="${axis}"${dataField} showAll="0"><items count="${
    items.length
  }">${items.join('')}</items></pivotField>`;
}

// Compact-form row lines: each distinct path through the row fields, in
// shared-item order, skipping children of collapsed items; then grand total.
function renderRowItems(model) {
  const {rows, records, cacheFields, collapsed} = model;
  const tree = new Map();
  records.forEach(record => {
    let node = tree;
    rows.forEach(field => {
      const x = cacheFields[field].sharedItems.indexOf(record[field]);
      if (!node.has(x)) node.set(x, new Map());
      node = node.get(x);
    });
  });

  const lines = [];
  const walk = (node, depth) => {
    [...node.keys()]
      .sort((a, b) => a - b)
      .forEach(x => {
        lines.push(`<i${depth ? ` r="${depth}"` : ''}>${x ? `<x v="${x}"/>` : '<x/>'}</i>`);
        if (!collapsed) walk(node.get(x), depth + 1);
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
