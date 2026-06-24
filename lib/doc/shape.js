const {parseRange} = require('./drawing-range');

class Shape {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    const m = {
      name: this.name,
      visible: this.visible,
      index: this.index,
      props: {
        type: this.props.type,
        rotation: this.props.rotation,
        horizontalFlip: this.props.horizontalFlip,
        verticalFlip: this.props.verticalFlip,
        fill: this.props.fill,
        outline: this.props.outline,
        textBody: this.props.textBody,
        bwMode: this.props.bwMode,
        extLstXml: this.props.extLstXml,
      },
      range: {
        tl: this.range.tl.model,
        br: this.range.br && this.range.br.model,
        ext: this.range.ext,
        editAs: this.range.editAs,
      },
      hyperlinks: this.hyperlinks,
    };
    if (this.txBox !== undefined) {
      m.txBox = this.txBox;
    }
    if (this.spLocks !== undefined) {
      m.spLocks = this.spLocks;
    }
    if (this.shapeType !== undefined) {
      m.shapeType = this.shapeType;
    }
    return m;
  }

  set model({name, visible, index, props, range, hyperlinks, txBox, spLocks, shapeType}) {
    const {name: propsName, visible: propsVisible, ...shapeProps} = props || {};
    this.name = name || propsName;
    const resolvedVisible = visible !== undefined ? visible : propsVisible;
    this.visible = resolvedVisible === undefined ? true : resolvedVisible;
    this.index = index;

    this.props = {type: shapeProps.type};
    if (shapeProps.rotation) {
      this.props.rotation = shapeProps.rotation;
    }
    if (shapeProps.horizontalFlip) {
      this.props.horizontalFlip = shapeProps.horizontalFlip;
    }
    if (shapeProps.verticalFlip) {
      this.props.verticalFlip = shapeProps.verticalFlip;
    }
    if (shapeProps.fill) {
      this.props.fill = shapeProps.fill;
    }
    if (shapeProps.outline) {
      this.props.outline = shapeProps.outline;
    }
    if (shapeProps.textBody) {
      this.props.textBody = parseAsTextBody(shapeProps.textBody);
    }
    if (shapeProps.bwMode) {
      this.props.bwMode = shapeProps.bwMode;
    }
    if (shapeProps.extLstXml) {
      this.props.extLstXml = shapeProps.extLstXml;
    }
    if (txBox !== undefined) {
      this.txBox = txBox;
    }
    if (spLocks !== undefined) {
      this.spLocks = spLocks;
    }
    if (shapeType !== undefined) {
      this.shapeType = shapeType;
    }
    this.range = parseRange(range, undefined, this.worksheet);
    this.hyperlinks = hyperlinks;
  }
}

function parseAsTextBody(input) {
  if (typeof input === 'string') {
    return {
      paragraphs: [parseAsParagraph(input)],
    };
  }
  if (Array.isArray(input)) {
    return {
      paragraphs: input.map(parseAsParagraph),
    };
  }
  const model = {
    paragraphs: input.paragraphs.map(parseAsParagraph),
  };
  if (input.vertAlign) {
    model.vertAlign = input.vertAlign;
  }
  // bodyPr 全属性を保持
  if (input.bodyPrAttrs) {
    model.bodyPrAttrs = input.bodyPrAttrs;
  }
  return model;
}

function parseAsParagraph(input) {
  if (typeof input === 'string') {
    return {
      runs: [parseAsRun(input)],
    };
  }
  if (Array.isArray(input)) {
    return {
      runs: input.map(parseAsRun),
    };
  }
  const model = {
    runs: input.runs.map(parseAsRun),
  };
  if (input.alignment) {
    model.alignment = input.alignment;
  }
  // pPr の追加属性を保持
  if (input.pPrAttrs) {
    model.pPrAttrs = input.pPrAttrs;
  }
  if (input.defRPr !== undefined) {
    model.defRPr = input.defRPr;
  }
  if (input.endParaRPr !== undefined) {
    model.endParaRPr = input.endParaRPr;
  }
  return model;
}

function parseAsRun(input) {
  if (typeof input === 'string') {
    return {
      text: input,
    };
  }
  const model = {
    text: input.text,
  };
  if (input.font) {
    model.font = input.font;
  }
  // rPr の詳細属性を保持
  if (input.rPrAttrs) {
    model.rPrAttrs = input.rPrAttrs;
  }
  if (input.rPrSolidFill) {
    model.rPrSolidFill = input.rPrSolidFill;
  }
  if (input.latinTypeface) {
    model.latinTypeface = input.latinTypeface;
  }
  if (input.eaTypeface) {
    model.eaTypeface = input.eaTypeface;
  }
  return model;
}

module.exports = Shape;
