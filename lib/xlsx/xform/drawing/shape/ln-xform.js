const BaseXform = require('../../base-xform');
const SolidFillXform = require('./solid-fill-xform');

// DocumentFormat.OpenXml.Drawing.Outline
class LnXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:solidFill': new SolidFillXform(),
    };
  }

  get tag() {
    return 'a:ln';
  }

  render(xmlStream, outline) {
    const attrs = {};
    if (outline.weight) {
      attrs.w = outline.weight * 12700;
    }
    if (outline.algn) {
      attrs.algn = outline.algn;
    }
    xmlStream.openNode(this.tag, attrs);
    const color = outline.color || outline;
    if (color) {
      this.map['a:solidFill'].render(xmlStream, color);
    }
    if (outline.dash) {
      xmlStream.leafNode('a:prstDash', {val: outline.dash});
    }
    if (outline.miter != null) {
      xmlStream.leafNode('a:miter', {lim: outline.miter});
    }
    if (outline.round) {
      xmlStream.leafNode('a:round');
    }
    if (outline.arrow) {
      if (outline.arrow.head) {
        xmlStream.leafNode('a:headEnd', {
          type: outline.arrow.head.type,
          w: outline.arrow.head.width,
          len: outline.arrow.head.length,
        });
      }
      if (outline.arrow.tail) {
        xmlStream.leafNode('a:tailEnd', {
          type: outline.arrow.tail.type,
          w: outline.arrow.tail.width,
          len: outline.arrow.tail.length,
        });
      }
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
        this.model = {};
        if (node.attributes.w) {
          this.model.weight = parseInt(node.attributes.w, 10) / 12700;
        }
        if (node.attributes.algn) {
          this.model.algn = node.attributes.algn;
        }
        break;
      case 'a:round':
        this.model.round = true;
        break;
      case 'a:miter':
        this.model.miter = node.attributes.lim != null ? parseInt(node.attributes.lim, 10) : 0;
        break;
      case 'a:prstDash':
        this.model.dash = node.attributes.val;
        break;
      case 'a:headEnd':
        this.model.arrow = this.model.arrow || {};
        this.model.arrow.head = {
          type: node.attributes.type,
          width: node.attributes.w,
          length: node.attributes.len,
        };
        break;
      case 'a:tailEnd':
        this.model.arrow = this.model.arrow || {};
        this.model.arrow.tail = {
          type: node.attributes.type,
          width: node.attributes.w,
          length: node.attributes.len,
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
        if (this.map['a:solidFill'].model) {
          this.model.color = this.map['a:solidFill'].model;
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = LnXform;
