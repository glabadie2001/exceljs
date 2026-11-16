const {xmlEncode} = require('../../../utils/utils');

// One cache item: <m/> blank, <n/> number, <s/> string.
function renderItem(value) {
  if (value === undefined) return '<m/>';
  if (typeof value === 'number') return `<n v="${value}"/>`;
  return `<s v="${xmlEncode(String(value))}"/>`;
}

class CacheField {
  constructor(field) {
    this.field = field;
  }

  render() {
    // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
    // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html
    const f = this.field;
    const attrs = [];
    // Defaults: containsSemiMixedTypes=1, containsString=1, others 0.
    if (!f.containsString && !f.containsBlank) attrs.push('containsSemiMixedTypes="0"');
    if (!f.containsString) attrs.push('containsString="0"');
    if (f.containsNumber && f.containsString) attrs.push('containsMixedTypes="1"');
    if (f.containsNumber) attrs.push('containsNumber="1"');
    if (f.containsInteger) attrs.push('containsInteger="1"');
    if (f.minValue !== undefined) attrs.push(`minValue="${f.minValue}" maxValue="${f.maxValue}"`);
    if (f.containsBlank) attrs.push('containsBlank="1"');

    const name = xmlEncode(String(f.name));
    if (f.sharedItems === null) {
      return `<cacheField name="${name}" numFmtId="0"><sharedItems ${attrs.join(
        ' '
      )}/></cacheField>`;
    }
    attrs.push(`count="${f.sharedItems.length}"`);
    return `<cacheField name="${name}" numFmtId="0"><sharedItems ${attrs.join(' ')}>${f.sharedItems
      .map(renderItem)
      .join('')}</sharedItems></cacheField>`;
  }
}

CacheField.renderItem = renderItem;

module.exports = CacheField;
