const {parseRange} = require('./drawing-range');

class Shape {
  constructor(worksheet, model) {
    this.worksheet = worksheet;
    this.model = model;
  }

  get model() {
    return {
      name: this.name,
      visible: this.visible,
      props: {
        type: this.props.type,
        rotation: this.props.rotation,
        horizontalFlip: this.props.horizontalFlip,
        verticalFlip: this.props.verticalFlip,
        fill: this.props.fill,
        outline: this.props.outline,
        textBody: this.props.textBody,
      },
      range: {
        tl: this.range.tl.model,
        br: this.range.br && this.range.br.model,
        ext: this.range.ext,
        editAs: this.range.editAs,
      },
      hyperlinks: this.hyperlinks,
    };
  }

  set model({name, visible, props, range, hyperlinks}) {
    const {name: propsName, visible: propsVisible, ...shapeProps} = props || {};
    this.name = name || propsName;
    const resolvedVisible = visible !== undefined ? visible : propsVisible;
    this.visible = resolvedVisible === undefined ? true : resolvedVisible;

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
  return model;
}

module.exports = Shape;
