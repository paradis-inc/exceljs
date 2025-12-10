const BaseXform = require('../base-xform');

// render the triangle in the cell for the comment
class VmlAnchorXform extends BaseXform {
  get tag() {
    return 'x:Anchor';
  }

  getAnchorRect(anchor) {
    const l = Math.floor(anchor.left);
    const lf = Math.floor((anchor.left - l) * 68);
    const t = Math.floor(anchor.top);
    const tf = Math.floor((anchor.top - t) * 18);
    const r = Math.floor(anchor.right);
    const rf = Math.floor((anchor.right - r) * 68);
    const b = Math.floor(anchor.bottom);
    const bf = Math.floor((anchor.bottom - b) * 18);
    return [l, lf, t, tf, r, rf, b, bf];
  }

  getDefaultRect(ref) {
    const l = ref.col;
    const lf = 6;
    const t = Math.max(ref.row - 2, 0);
    const tf = 14;
    const r = l + 2;
    const rf = 2;
    const b = t + 4;
    const bf = 16;
    return [l, lf, t, tf, r, rf, b, bf];
  }

  render(xmlStream, model) {
    let rect;
    if (model.anchor && model.anchor.raw && model.anchor.raw.length === 8) {
      rect = model.anchor.raw;
    } else if (model.anchor) {
      rect = this.getAnchorRect(model.anchor);
    } else {
      rect = this.getDefaultRect(model.refAddress);
    }

    xmlStream.leafNode('x:Anchor', null, rect.join(', '));
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.text = '';
        this.model = undefined;
        return true;
      default:
        return false;
    }
  }

  parseText(text) {
    this.text = text;
  }

  parseClose() {
    if (this.text) {
      const parts = this.text
        .split(',')
        .map(p => Number(p.trim()))
        .filter(n => Number.isFinite(n));

      if (parts.length === 8) {
        const [l, lf, t, tf, r, rf, b, bf] = parts;
        this.model = {
          raw: parts,
          left: l + lf / 68,
          top: t + tf / 18,
          right: r + rf / 68,
          bottom: b + bf / 18,
        };
      }
    }
    return false;
  }
}

module.exports = VmlAnchorXform;
