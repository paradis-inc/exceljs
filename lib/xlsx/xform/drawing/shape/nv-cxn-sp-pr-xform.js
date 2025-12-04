const BaseXform = require('../../base-xform');
const StaticXform = require('../../static-xform');
const CNvPrXform = require('../c-nv-pr-xform');

class NvCxnSpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:cNvPr': new CNvPrXform(false),
      'xdr:cNvCxnSpPr': new StaticXform({tag: 'xdr:cNvCxnSpPr'}),
    };
  }

  get tag() {
    return 'xdr:nvCxnSpPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    this.map['xdr:cNvPr'].render(xmlStream, shape);
    this.map['xdr:cNvCxnSpPr'].render(xmlStream, {});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
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
      case this.tag:
        const cNvPrModel = this.map['xdr:cNvPr'].model || {};
        this.model = {
          name: cNvPrModel.name,
          visible: cNvPrModel.visible,
          hyperlinks: cNvPrModel.hyperlinks,
        };
        return false;
      default:
        return true;
    }
  }
}

module.exports = NvCxnSpPrXform;
