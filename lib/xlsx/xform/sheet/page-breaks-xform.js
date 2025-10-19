const BaseXform = require('../base-xform');

class PageBreaksXform extends BaseXform {
  get tag() {
    return 'brk';
  }

  render(xmlStream, model) {
    const attributes = {
      id: model.id,
    };
    if (model.max !== undefined) {
      attributes.max = model.max;
    }
    if (model.man !== undefined) {
      attributes.man = model.man ? '1' : '0';
    }
    xmlStream.leafNode('brk', attributes);
  }

  parseOpen(node) {
    if (node.name === 'brk') {
      this.model = {
        id: parseInt(node.attributes.id, 10),
        max: node.attributes.max ? parseInt(node.attributes.max, 10) : undefined,
        man: node.attributes.man === '1',
      };
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = PageBreaksXform;
