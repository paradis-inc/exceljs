const SpXform = require('./sp-xform');
const NvCxnSpPrXform = require('./nv-cxn-sp-pr-xform');

class CxnSpXform extends SpXform {
  constructor() {
    super({tag: 'xdr:cxnSp', nvTag: 'xdr:nvCxnSpPr', nvXform: new NvCxnSpPrXform()});
  }

  render(xmlStream, shape) {
    // xdr:cxnSp は macro/textlink 属性不要、xdr:txBody も不要
    // xdr:style は a:lnRef（線の色）を保持するため shape.style があれば出力する
    xmlStream.openNode(this.tag);
    this.map[this.nvTag].render(xmlStream, shape);
    this.map['xdr:spPr'].render(xmlStream, shape.props);
    if (shape.style) {
      this.map['xdr:style'].render(xmlStream, shape.style);
    }
    xmlStream.closeNode();
  }
}

module.exports = CxnSpXform;
