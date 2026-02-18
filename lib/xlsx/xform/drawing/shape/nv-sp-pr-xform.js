const BaseXform = require('../../base-xform');
const CNvPrXform = require('../c-nv-pr-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties
class NvSpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:cNvPr': new CNvPrXform(false),
    };
  }

  get tag() {
    return 'xdr:nvSpPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    this.map['xdr:cNvPr'].render(xmlStream, shape);
    const cNvSpPrAttrs = shape.txBox ? {txBox: '1'} : undefined;
    // spLocks が保存されている場合は子要素として出力、なければ leafNode
    if (shape.spLocks && shape.spLocks.length > 0) {
      xmlStream.openNode('xdr:cNvSpPr', cNvSpPrAttrs);
      for (const lock of shape.spLocks) {
        xmlStream.leafNode('a:spLocks', lock);
      }
      xmlStream.closeNode();
    } else {
      xmlStream.leafNode('xdr:cNvSpPr', cNvSpPrAttrs);
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.cNvSpPrAttrs = null;
        this._spLocks = [];
        this._inCNvSpPr = false;
        break;
      case 'xdr:cNvSpPr':
        this.cNvSpPrAttrs = node.attributes;
        this._inCNvSpPr = true;
        break;
      case 'a:spLocks':
        if (this._inCNvSpPr) {
          this._spLocks.push(Object.assign({}, node.attributes));
        } else {
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
        }
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'xdr:cNvSpPr':
        this._inCNvSpPr = false;
        return true;
      case this.tag: {
        const cNvPrModel = this.map['xdr:cNvPr'].model || {};
        this.model = {
          name: cNvPrModel.name,
          visible: cNvPrModel.visible,
          hyperlinks: cNvPrModel.hyperlinks,
        };
        if (this.cNvSpPrAttrs && this.cNvSpPrAttrs.txBox === '1') {
          this.model.txBox = true;
        }
        if (this._spLocks && this._spLocks.length > 0) {
          this.model.spLocks = this._spLocks;
        }
        if (cNvPrModel.extLstXml !== undefined) {
          this.model.extLstXml = cNvPrModel.extLstXml;
        }
        return false;
      }
      default:
        return true;
    }
  }
}

module.exports = NvSpPrXform;
