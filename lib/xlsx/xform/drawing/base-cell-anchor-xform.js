const BaseXform = require('../base-xform');
const xmlUtils = require('../../../utils/utils');


class BaseCellAnchorXform extends BaseXform {
  parseOpen(node) {
    // raw XML 収集モード中（xdr:grpSp など未知タグの処理）
    if (this._rawDepth > 0) {
      this._rawDepth++;
      const attrsStr = Object.entries(node.attributes || {})
        .map(([k, v]) => ` ${k}="${xmlUtils.xmlEncode(String(v))}"`)
        .join('');
      this._rawParts.push(`<${node.name}${attrsStr}>`);
      return true;
    }

    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this._rawDepth = 0;
        this._rawParts = null;
        this.model = {
          range: {
            editAs: node.attributes.editAs || undefined,
          },
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        } else {
          // 未知タグ（xdr:grpSp 等）→ raw XML として収集開始
          this._rawDepth = 1;
          this._rawParts = [];
          const attrsStr = Object.entries(node.attributes || {})
            .map(([k, v]) => ` ${k}="${xmlUtils.xmlEncode(String(v))}"`)
            .join('');
          this._rawParts.push(`<${node.name}${attrsStr}>`);
        }
        break;
    }
    return true;
  }

  parseText(text) {
    if (this._rawDepth > 0) {
      if (text) this._rawParts.push(text);
      return;
    }
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    // raw XML 収集モード中
    if (this._rawDepth > 0) {
      this._rawParts.push(`</${name}>`);
      this._rawDepth--;
      if (this._rawDepth === 0) {
        if (!this.model.rawXmls) this.model.rawXmls = [];
        this.model.rawXmls.push(this._rawParts.join(''));
        this._rawParts = null;
      }
      return true;
    }

    // サブクラスで実装
    return true;
  }

  reconcilePicture(model, options) {
    if (model && model.rId) {
      const rel = options.rels[model.rId];
      const match = rel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
      if (match) {
        const name = match[1];
        const mediaId = options.mediaIndex[name];
        return options.media[mediaId];
      }
    }
    return undefined;
  }
}

module.exports = BaseCellAnchorXform;
