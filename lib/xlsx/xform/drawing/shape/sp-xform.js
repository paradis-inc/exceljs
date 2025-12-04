const BaseXform = require('../../base-xform');

const NvSpPrXform = require('./nv-sp-pr-xform');
const SpPrXform = require('./sp-pr-xform');
const StyleXform = require('./style-xform');
const TxBodyXform = require('./tx-body-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape
class SpXform extends BaseXform {
  constructor(options = {}) {
    super();

    this.tagName = options.tag || 'xdr:sp';
    this.nvTag = options.nvTag || 'xdr:nvSpPr';
    const nvXform = options.nvXform || new NvSpPrXform();

    this.map = {
      [this.nvTag]: nvXform,
      'xdr:spPr': new SpPrXform(),
      'xdr:style': new StyleXform(),
      'xdr:txBody': new TxBodyXform(),
    };
  }

  get tag() {
    return this.tagName;
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag, {macro: '', textlink: ''});

    this.map[this.nvTag].render(xmlStream, shape);
    this.map['xdr:spPr'].render(xmlStream, shape.props);
    this.map['xdr:style'].render(xmlStream, shape.props);
    if (shape.props.textBody) {
      this.map['xdr:txBody'].render(xmlStream, shape.props.textBody);
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      this.model = {props: {}};
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

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        if (this.map['xdr:style'].model) {
          this.model.props = {
            ...(this.map['xdr:style'].model.fill ? {fill: this.map['xdr:style'].model.fill} : {}),
            ...(this.map['xdr:style'].model.outline ? {outline: this.map['xdr:style'].model.outline} : {}),
          };
        }
        if (this.map['xdr:spPr'].model) {
          this.model.props = {
            ...this.model.props,
            ...this.map['xdr:spPr'].model,
          };
        }
        if (this.map['xdr:txBody'].model) {
          this.model.props.textBody = this.map['xdr:txBody'].model;
        }
        if (this.map['xdr:spPr'].noFill) {
          delete this.model.props.fill;
        }
        if (this.map[this.nvTag].model) {
          const nvModel = this.map[this.nvTag].model;
          if (nvModel.hyperlinks) {
            this.model.hyperlinks = nvModel.hyperlinks;
          }
          if (nvModel.name !== undefined) {
            this.model.name = nvModel.name;
          }
          if (nvModel.visible !== undefined) {
            this.model.visible = nvModel.visible;
          }
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpXform;
