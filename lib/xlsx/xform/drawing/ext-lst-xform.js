const BaseXform = require('../base-xform');
const xmlUtils = require('../../../utils/utils');

function xmlEncodeAttr(str) {
  return xmlUtils.xmlEncode(String(str));
}

class ExtLstXform extends BaseXform {
  get tag() {
    return 'a:extLst';
  }

  render(xmlStream, model) {
    // raw XML が保存されている場合はそのまま出力
    if (model && model.extLstXml) {
      xmlStream.writeXml(model.extLstXml);
      return;
    }
    // raw XML がない場合はデフォルトの creationId を出力
    xmlStream.openNode(this.tag);
    xmlStream.openNode('a:ext', {
      uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}',
    });
    xmlStream.leafNode('a16:creationId', {
      'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
      id: '{00000000-0008-0000-0000-000002000000}',
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this._depth = 1;
        this._parts = ['<a:extLst>'];
        return true;
      default:
        if (this._depth > 0) {
          this._depth++;
          const attrsStr = Object.entries(node.attributes || {})
            .map(([k, v]) => ` ${k}="${xmlEncodeAttr(v)}"`)
            .join('');
          this._parts.push(`<${node.name}${attrsStr}>`);
        }
        return true;
    }
  }

  parseText(text) {
    if (this._depth > 0 && text) {
      this._parts.push(text);
    }
  }

  parseClose(name) {
    if (name === this.tag) {
      this._depth = 0;
      this._parts.push('</a:extLst>');
      this.model = {extLstXml: this._parts.join('')};
      this._parts = null;
      return false;
    }
    if (this._depth > 0) {
      this._depth--;
      this._parts.push(`</${name}>`);
    }
    return true;
  }
}

module.exports = ExtLstXform;
