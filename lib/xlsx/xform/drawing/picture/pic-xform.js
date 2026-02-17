const BaseXform = require('../../base-xform');
const XfrmXform = require('../xfrm-xform');

const BlipFillXform = require('../blip-fill-xform');
const NvPicPrXform = require('./nv-pic-pr-xform');

class PicXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:nvPicPr': new NvPicPrXform(),
      'xdr:blipFill': new BlipFillXform(),
      'xdr:spPr': new PicSpPrXform(),
    };
  }

  get tag() {
    return 'xdr:pic';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvPicPr'].render(xmlStream, model);
    this.map['xdr:blipFill'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
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
        this.mergeModel(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

// picture 用の xdr:spPr: a:xfrm の off/ext を parse/render する
class PicSpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
    };
  }

  get tag() {
    return 'xdr:spPr';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    // spPr.xfrm が存在しない場合は空オブジェクト（off/ext が 0 になる）でフォールバック
    const xfrmModel = (model && model.spPr && model.spPr.xfrm) || {};
    this.map['a:xfrm'].render(xmlStream, xfrmModel);

    xmlStream.openNode('a:prstGeom', {prst: 'rect'});
    xmlStream.leafNode('a:avLst');
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = {spPr: {}};
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
        if (this.parser === this.map['a:xfrm'] && this.parser.model) {
          this.model.spPr.xfrm = this.parser.model;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = PicXform;
