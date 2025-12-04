const SpXform = require('./sp-xform');
const NvCxnSpPrXform = require('./nv-cxn-sp-pr-xform');

class CxnSpXform extends SpXform {
  constructor() {
    super({tag: 'xdr:cxnSp', nvTag: 'xdr:nvCxnSpPr', nvXform: new NvCxnSpPrXform()});
  }
}

module.exports = CxnSpXform;
