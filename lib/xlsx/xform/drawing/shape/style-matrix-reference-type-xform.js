const BaseXform = require('../../base-xform');

// DocumentFormat.OpenXml.Drawing.StyleMatrixReferenceType
class StyleMatrixReferenceTypeXform extends BaseXform {
  constructor(tagName) {
    super();

    this.map = {};
    this.tagName = tagName;
  }

  get tag() {
    return this.tagName;
  }

  idx(defaultIdx) {
    if (this.model && typeof this.model.idx === 'number') {
      return this.model.idx;
    }
    switch (this.tagName) {
      case 'a:lnRef':
        return 2;
      case 'a:fillRef':
        return 1;
      default:
        return defaultIdx || 0;
    }
  }

  render(xmlStream, shape) {
    const idx =
      (shape && shape.outline && typeof shape.outline.idx === 'number' && shape.outline.idx) ||
      (shape && shape.fill && typeof shape.fill.idx === 'number' && shape.fill.idx) ||
      this.idx();
    const colorSource =
      this.tagName === 'a:lnRef'
        ? (shape && (shape.outline?.color || shape.outline))
        : this.tagName === 'a:fillRef'
          ? (shape && (shape.fill?.color || shape.fill))
          : null;

    xmlStream.openNode(this.tag, {idx});
    if (colorSource) {
      if (colorSource.theme) {
        xmlStream.leafNode('a:schemeClr', {val: colorSource.theme});
      } else if (colorSource.rgb) {
        xmlStream.leafNode('a:srgbClr', {val: colorSource.rgb});
      } else {
        // fallback to previous behavior
        xmlStream.leafNode('a:schemeClr', {val: 'accent1'});
      }
    } else {
      // default behavior when no color found
      xmlStream.leafNode('a:schemeClr', {val: 'accent1'});
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
        this.model = { idx: node.attributes && node.attributes.idx ? Number(node.attributes.idx) : undefined };
        break;
      case 'a:schemeClr':
        this.model.theme = node.attributes.val;
        break;
      case 'a:srgbClr':
        this.model.rgb = node.attributes.val;
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
        return false;
      default:
        return true;
    }
  }
}

module.exports = StyleMatrixReferenceTypeXform;
