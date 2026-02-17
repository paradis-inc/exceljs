const BaseXform = require('../../base-xform');
const ParagraphXform = require('./p-xform');

// bodyPr の保持する属性一覧
const BODY_PR_ATTRS = [
  'vertOverflow', 'horzOverflow', 'wrap', 'lIns', 'tIns', 'rIns', 'bIns',
  'anchor', 'anchorCtr', 'upright', 'rtlCol', 'rot', 'spcFirstLastPara',
  'vert', 'numCol', 'spcCol', 'forceAA', 'compatLnSpc',
];

// DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody
class TxBodyXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:p': new ParagraphXform(),
    };
  }

  get tag() {
    return 'xdr:txBody';
  }

  render(xmlStream, textBody) {
    xmlStream.openNode(this.tag);

    // bodyPr: 保存した属性をそのまま出力。保存がなければデフォルト
    const bodyPrAttrs = textBody.bodyPrAttrs || {
      vertOverflow: 'clip',
      horzOverflow: 'clip',
      rtlCol: '0',
    };
    // anchor は vertAlign としても保存されている場合の後方互換
    if (!bodyPrAttrs.anchor && textBody.vertAlign) {
      bodyPrAttrs.anchor = textBody.vertAlign;
    }
    xmlStream.leafNode('a:bodyPr', bodyPrAttrs);

    xmlStream.leafNode('a:lstStyle');
    textBody.paragraphs.forEach(p => {
      this.map['a:p'].render(xmlStream, p);
    });
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {paragraphs: []};
        break;
      case 'a:bodyPr': {
        // 全属性を bodyPrAttrs として保存
        const attrs = {};
        for (const key of BODY_PR_ATTRS) {
          if (node.attributes[key] !== undefined) {
            attrs[key] = node.attributes[key];
          }
        }
        // その他の属性も保存
        for (const key of Object.keys(node.attributes)) {
          if (!(key in attrs)) {
            attrs[key] = node.attributes[key];
          }
        }
        this.model.bodyPrAttrs = attrs;
        // 後方互換のため vertAlign も保存
        if (node.attributes.anchor) {
          this.model.vertAlign = node.attributes.anchor;
        }
        break;
      }
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
        if (name === 'a:p') {
          this.model.paragraphs.push(this.parser.model);
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

module.exports = TxBodyXform;
