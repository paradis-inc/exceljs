const BaseXform = require('../../base-xform');
const SolidFillXform = require('./solid-fill-xform');

// rPr で保持する属性一覧
const RPR_ATTRS = ['lang', 'altLang', 'sz', 'b', 'i', 'u', 'strike', 'baseline', 'dirty', 'smtId', 'kumimoji', 'spc', 'normalizeH', 'noProof', 'err'];

// DocumentFormat.OpenXml.Drawing.Run
class RunXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:solidFill': new SolidFillXform(),
    };
  }

  get tag() {
    return 'a:r';
  }

  render(xmlStream, run) {
    xmlStream.openNode(this.tag);
    // rPr: 保存した属性をそのまま出力
    const rPrAttrs = {};
    if (run.rPrAttrs) {
      Object.assign(rPrAttrs, run.rPrAttrs);
    } else if (run.font) {
      // 後方互換: font オブジェクトから属性を生成
      if (run.font.size) rPrAttrs.sz = run.font.size * 100;
      if (run.font.bold) rPrAttrs.b = 1;
      if (run.font.italic) rPrAttrs.i = 1;
      if (run.font.underline) rPrAttrs.u = run.font.underline;
    }
    xmlStream.openNode('a:rPr', rPrAttrs);
    if (run.font && run.font.color) {
      this.map['a:solidFill'].render(xmlStream, run.font.color);
    } else if (run.rPrSolidFill) {
      this.map['a:solidFill'].render(xmlStream, run.rPrSolidFill);
    }
    // a:latin / a:ea フォント指定
    if (run.latinTypeface) {
      xmlStream.leafNode('a:latin', {typeface: run.latinTypeface});
    }
    if (run.eaTypeface) {
      xmlStream.leafNode('a:ea', {typeface: run.eaTypeface});
    }
    xmlStream.closeNode();
    xmlStream.leafNode('a:t', undefined, run.text);
    xmlStream.closeNode();
  }

  // endParaRPr 用レンダラー（runs がない段落の終端用）
  renderEndParaRPr(xmlStream, endParaRPr) {
    if (!endParaRPr) return;
    xmlStream.openNode('a:endParaRPr', endParaRPr.attrs || {});
    if (endParaRPr.solidFill) {
      this.map['a:solidFill'].render(xmlStream, endParaRPr.solidFill);
    }
    if (endParaRPr.latinTypeface) {
      xmlStream.leafNode('a:latin', {typeface: endParaRPr.latinTypeface});
    }
    if (endParaRPr.eaTypeface) {
      xmlStream.leafNode('a:ea', {typeface: endParaRPr.eaTypeface});
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
        this.model = {text: '', font: {}};
        this.parsingText = false;
        this._inRPr = false;
        break;
      case 'a:rPr': {
        this._inRPr = true;
        // 全属性を rPrAttrs として保存
        this.model.rPrAttrs = Object.assign({}, node.attributes);
        // 後方互換のため font にも保存
        if (node.attributes.sz) {
          this.model.font.size = parseInt(node.attributes.sz, 10) / 100;
        }
        if (node.attributes.b) {
          this.model.font.bold = node.attributes.b === '1';
        }
        if (node.attributes.i) {
          this.model.font.italic = node.attributes.i === '1';
        }
        if (node.attributes.u) {
          this.model.font.underline = node.attributes.u;
        }
        break;
      }
      case 'a:latin':
        if (this._inRPr) {
          this.model.latinTypeface = node.attributes.typeface;
        }
        break;
      case 'a:ea':
        if (this._inRPr) {
          this.model.eaTypeface = node.attributes.typeface;
        }
        break;
      case 'a:t':
        this.parsingText = true;
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
    if (this.parsingText) {
      this.model.text = text.replace(/_x([0-9A-F]{4})_/g, ($0, $1) => String.fromCharCode(parseInt($1, 16)));
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        if (name === 'a:solidFill' && this._inRPr) {
          this.model.rPrSolidFill = this.map['a:solidFill'].model;
          this.model.font.color = this.map['a:solidFill'].model;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'a:rPr':
        this._inRPr = false;
        return true;
      case 'a:latin':
      case 'a:ea':
        return true;
      case this.tag:
        if (this.map['a:solidFill'].model && !this.model.rPrSolidFill) {
          this.model.font.color = this.map['a:solidFill'].model;
        }
        return false;
      case 'a:t':
        this.parsingText = false;
        return true;
      default:
        return true;
    }
  }
}

module.exports = RunXform;
