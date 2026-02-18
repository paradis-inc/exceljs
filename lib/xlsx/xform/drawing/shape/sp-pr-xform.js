const BaseXform = require('../../base-xform');
const XfrmXform = require('../xfrm-xform');
const PrstGeomXform = require('./prst-geom-xform');
const SolidFillXform = require('./solid-fill-xform');
const LnXform = require('./ln-xform');
const xmlUtils = require('../../../../utils/utils');


// DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties
class SpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
      'a:prstGeom': new PrstGeomXform(),
      'a:solidFill': new SolidFillXform(),
      'a:ln': new LnXform(),
    };
  }

  get tag() {
    return 'xdr:spPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag, shape.bwMode ? {bwMode: shape.bwMode} : undefined);
    this.map['a:xfrm'].render(xmlStream, shape);
    this.map['a:prstGeom'].render(xmlStream, shape);
    if (shape.fill?._fromStyle) {
      // xdr:style > a:fillRef に任せる。xdr:spPr には fill 要素を出力しない
    } else if (shape.fill && shape.fill.type === 'solid') {
      this.map['a:solidFill'].render(xmlStream, shape.fill.color);
    } else {
      xmlStream.leafNode('a:noFill');
    }
    const shouldRenderLn = shape.outline && !(
      shape.outline._fromStyle &&
      !shape.outline.weight &&
      !shape.outline.dash &&
      !shape.outline.arrow
    );
    if (shouldRenderLn) {
      this.map['a:ln'].render(xmlStream, shape.outline);
    }
    // a:extLst の raw XML があれば出力
    if (shape.extLstXml) {
      xmlStream.writeXml(shape.extLstXml);
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
        this.noFill = false;
        this._extLstDepth = 0;
        this._extLstParts = null;
        if (node.attributes.bwMode) {
          this.model.bwMode = node.attributes.bwMode;
        }
        break;
      case 'a:extLst':
        // extLst を raw XML として収集開始
        this._extLstDepth = 1;
        this._extLstParts = ['<a:extLst>'];
        break;
      case 'a:noFill':
        if (this._extLstDepth > 0) {
          // extLst 内の noFill は後で closeNode が来るので開きタグだけ出す
          this._extLstDepth++;
          this._extLstParts.push('<a:noFill');
          this._extLstParts.push('>');
        } else {
          this.noFill = true;
        }
        break;
      default:
        if (this._extLstDepth > 0) {
          this._extLstDepth++;
          // 開きタグを収集
          const attrsStr = Object.entries(node.attributes || {})
            .map(([k, v]) => ` ${k}="${xmlUtils.xmlEncode(String(v))}"`)
            .join('');
          this._extLstParts.push(`<${node.name}${attrsStr}>`);
        } else {
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
        }
        break;
    }
    return true;
  }

  parseText(text) {
    if (this._extLstDepth > 0) {
      if (text) this._extLstParts.push(text);
    }
  }

  parseClose(name) {
    if (this._extLstDepth > 0) {
      if (name === 'a:extLst') {
        this._extLstDepth = 0;
        this._extLstParts.push('</a:extLst>');
        this.model.extLstXml = this._extLstParts.join('');
        this._extLstParts = null;
      } else {
        this._extLstDepth--;
        this._extLstParts.push(`</${name}>`);
      }
      return true;
    }
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        if (this.map['a:prstGeom'].model) {
          this.model.type = this.map['a:prstGeom'].model.type;
        }
        if (this.map['a:solidFill'].model) {
          this.model.fill = {
            type: 'solid',
            color: this.map['a:solidFill'].model,
          };
        }
        if (this.map['a:ln'].model) {
          this.model.outline = this.map['a:ln'].model;
        }
        if (this.map['a:xfrm'].model) {
          this.mergeModel(this.map['a:xfrm'].model);
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpPrXform;
