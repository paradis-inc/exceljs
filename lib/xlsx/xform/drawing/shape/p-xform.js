const BaseXform = require('../../base-xform');
const RunXform = require('./r-xform');
const SolidFillXform = require('./solid-fill-xform');

// DocumentFormat.OpenXml.Drawing.Paragraph
class ParagraphXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:r': new RunXform(),
      'a:solidFill': new SolidFillXform(),
    };
  }

  get tag() {
    return 'a:p';
  }

  render(xmlStream, paragraph) {
    xmlStream.openNode('a:p');
    // pPr: 保存した属性をそのまま出力
    const pPrAttrs = {};
    if (paragraph.alignment) {
      pPrAttrs.algn = paragraph.alignment;
    }
    if (paragraph.pPrAttrs) {
      Object.assign(pPrAttrs, paragraph.pPrAttrs);
    }
    if (paragraph.defRPr !== undefined) {
      xmlStream.openNode('a:pPr', pPrAttrs);
      xmlStream.leafNode('a:defRPr', paragraph.defRPr);
      xmlStream.closeNode();
    } else if (Object.keys(pPrAttrs).length > 0) {
      xmlStream.leafNode('a:pPr', pPrAttrs);
    } else {
      xmlStream.leafNode('a:pPr', pPrAttrs);
    }

    // runs を出力
    paragraph.runs.forEach(r => {
      this.map['a:r'].render(xmlStream, r);
    });

    // endParaRPr があれば出力
    if (paragraph.endParaRPr !== undefined) {
      this.map['a:r'].renderEndParaRPr(xmlStream, paragraph.endParaRPr);
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
        this.model = {runs: []};
        this._inPPr = false;
        this._inEndParaRPr = false;
        this._endParaRPr = null;
        break;
      case 'a:pPr': {
        this._inPPr = true;
        if (node.attributes.algn) {
          this.model.alignment = node.attributes.algn;
        }
        // 全属性を pPrAttrs として保存（algn 以外も）
        const attrs = {};
        for (const key of Object.keys(node.attributes)) {
          if (key !== 'algn') {
            attrs[key] = node.attributes[key];
          }
        }
        if (Object.keys(attrs).length > 0) {
          this.model.pPrAttrs = attrs;
        }
        break;
      }
      case 'a:defRPr':
        if (this._inPPr) {
          // defRPr の属性を保存
          this.model.defRPr = Object.assign({}, node.attributes);
        }
        break;
      case 'a:endParaRPr':
        // endParaRPr を保存開始
        this._inEndParaRPr = true;
        this._endParaRPr = {attrs: Object.assign({}, node.attributes)};
        break;
      case 'a:solidFill':
        if (this._inEndParaRPr) {
          this.parser = this.map['a:solidFill'];
          this.parser.parseOpen(node);
        } else {
          this.parser = this.map['a:solidFill'];
          this.parser.parseOpen(node);
        }
        break;
      case 'a:latin':
        if (this._inEndParaRPr) {
          this._endParaRPr.latinTypeface = node.attributes.typeface;
        }
        break;
      case 'a:ea':
        if (this._inEndParaRPr) {
          this._endParaRPr.eaTypeface = node.attributes.typeface;
        }
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
        if (name === 'a:r') {
          this.model.runs.push(this.parser.model);
        } else if (name === 'a:solidFill' && this._inEndParaRPr && this._endParaRPr) {
          this._endParaRPr.solidFill = this.map['a:solidFill'].model;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'a:pPr':
        this._inPPr = false;
        return true;
      case 'a:defRPr':
        return true;
      case 'a:latin':
      case 'a:ea':
        return true;
      case 'a:endParaRPr':
        this._inEndParaRPr = false;
        this.model.endParaRPr = this._endParaRPr;
        return true;
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = ParagraphXform;
