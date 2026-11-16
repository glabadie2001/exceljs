const XmlStream = require('../../../utils/xml-stream');

const BaseXform = require('../base-xform');
const CacheField = require('./cache-field');

class PivotCacheRecordsXform extends BaseXform {
  constructor() {
    super();

    this.map = {};
  }

  prepare(model) {
    // TK
  }

  get tag() {
    // http://www.datypic.com/sc/ooxml/e-ssml_pivotCacheRecords.html
    return 'pivotCacheRecords';
  }

  render(xmlStream, model) {
    const {records, cacheFields} = model;

    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES,
      count: records.length,
    });
    // PivotCache Record: http://www.datypic.com/sc/ooxml/e-ssml_r-1.html
    // Axis fields point into sharedItems (<x/>); others hold the value.
    xmlStream.writeXml(
      records
        .map(
          record =>
            `<r>${record
              .map((value, i) => {
                const {sharedItems} = cacheFields[i];
                return sharedItems
                  ? `<x v="${sharedItems.indexOf(value)}"/>`
                  : CacheField.renderItem(value);
              })
              .join('')}</r>`
        )
        .join('')
    );
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

PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES = {
  xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'mc:Ignorable': 'xr',
  'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
};

module.exports = PivotCacheRecordsXform;
