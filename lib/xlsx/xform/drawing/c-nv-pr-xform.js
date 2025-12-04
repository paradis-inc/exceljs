const BaseXform = require('../base-xform');
const HlickClickXform = require('./hlink-click-xform');
const ExtLstXform = require('./ext-lst-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
class CNvPrXform extends BaseXform {
  constructor(isPicture) {
    super();

    this.isPicture = isPicture;

    this.map = {
      'a:hlinkClick': new HlickClickXform(),
      'a:extLst': new ExtLstXform(),
    };
  }

  get tag() {
    return 'xdr:cNvPr';
  }

  render(xmlStream, model) {
    const name = model.name || `${this.isPicture ? 'Picture' : 'Shape'} ${model.index}`;
    const attributes = {
      id: model.index,
      name,
    };
    if (model.visible === false) {
      attributes.hidden = '1';
    }

    xmlStream.openNode(this.tag, attributes);
    this.map['a:hlinkClick'].render(xmlStream, model);
    this.map['a:extLst'].render(xmlStream, model);
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
        this.model = {
          index: BaseXform.toIntValue(node.attributes.id),
          name: node.attributes.name,
          visible: node.attributes.hidden ? node.attributes.hidden !== '1' : true,
        };
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
        const hyperlinkModel = this.map['a:hlinkClick'].model;
        if (hyperlinkModel && hyperlinkModel.hyperlinks) {
          this.model = this.model || {};
          this.model.hyperlinks = hyperlinkModel.hyperlinks;
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = CNvPrXform;
