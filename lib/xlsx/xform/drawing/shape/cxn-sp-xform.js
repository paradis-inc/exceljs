const SpXform = require('./sp-xform');
const NvCxnSpPrXform = require('./nv-cxn-sp-pr-xform');

class CxnSpXform extends SpXform {
  constructor() {
    super({tag: 'xdr:cxnSp', nvTag: 'xdr:nvCxnSpPr', nvXform: new NvCxnSpPrXform()});
  }

  render(xmlStream, shape) {
    // xdr:cxnSp は macro/textlink 属性不要、xdr:style/xdr:txBody も不要
    xmlStream.openNode(this.tag);
    this.map[this.nvTag].render(xmlStream, shape);
    this.map['xdr:spPr'].render(xmlStream, shape.props);
    xmlStream.closeNode();
  }
}

module.exports = CxnSpXform;
