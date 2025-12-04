const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
const WITH_LINE_FIXTURE = './test/data/withline.xlsx';

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Shapes', () => {
    it('stores shape', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      let wb2;
      let ws2;

      ws.addShape(
        {
          type: 'line',
          fill: {type: 'solid', color: {theme: 'accent6'}},
          outline: {
            weight: 30000,
            color: {theme: 'accent1'},
            arrow: {
              head: {type: 'triangle', width: 'lg', length: 'med'},
            },
          },
        },
        'B2:D6'
      );

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('sheet');
          expect(ws2).to.not.be.undefined();
          const shapes = ws2.getShapes();
          expect(shapes.length).to.equal(1);

          const shape = shapes[0];
          expect(shape.props.type).to.equal('line');
          expect(shape.props.fill).to.deep.equal({
            type: 'solid',
            color: {theme: 'accent6'},
          });
          expect(shape.props.outline).to.deep.equal({
            weight: 30000,
            color: {theme: 'accent1'},
            arrow: {
              head: {type: 'triangle', width: 'lg', length: 'med'},
            },
          });
        });
    });

    it('stores shape with oneCell', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      let wb2;
      let ws2;

      ws.addShape(
        {
          type: 'rect',
          rotation: 180,
          horizontalFlip: true,
          fill: {type: 'solid', color: {rgb: 'AABBCC'}},
        },
        {
          tl: {col: 0.1125, row: 0.4},
          br: {col: 2.101046875, row: 3.4},
          editAs: 'oneCell',
        }
      );

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('sheet');
          expect(ws2).to.not.be.undefined();
          const shapes = ws2.getShapes();
          expect(shapes.length).to.equal(1);

          const shape = shapes[0];
          expect(shape.range.editAs).to.equal('oneCell');
          expect(shape.props.type).to.equal('rect');
          expect(shape.props.rotation).to.equal(180);
          expect(shape.props.horizontalFlip).to.equal(true);
          expect(shape.props.fill).to.deep.equal({
            type: 'solid',
            color: {rgb: 'AABBCC'},
          });
        });
    });

    it('stores shape name and visibility', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      let wb2;
      let ws2;

      ws.addShape(
        {
          type: 'line',
          name: 'Named Line',
          visible: false,
          fill: {type: 'solid', color: {theme: 'accent6'}},
        },
        'B2:D6'
      );

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          ws2 = wb2.getWorksheet('sheet');
          expect(ws2).to.not.be.undefined();
          const shapes = ws2.getShapes();
          expect(shapes.length).to.equal(1);
          const shape = shapes[0];
          expect(shape.name).to.equal('Named Line');
          expect(shape.visible).to.equal(false);
        });
    });

    it('reads existing shape names and toggles visibility', () => {
      const wb = new ExcelJS.Workbook();
      let wb2;

      return wb.xlsx
        .readFile(WITH_LINE_FIXTURE)
        .then(() => {
          const ws = wb.getWorksheet('Sheet1');
          expect(ws).to.not.be.undefined();
          const shapes = ws.getShapes();
          expect(shapes.length).to.equal(1);
          expect(shapes[0].name).to.equal('直線コネクタ 2');
          expect(shapes[0].visible).to.equal(true);

          shapes[0].visible = false;
          return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          const ws = wb2.getWorksheet('Sheet1');
          const shapes = ws.getShapes();
          expect(shapes.length).to.equal(1);
          expect(shapes[0].name).to.equal('直線コネクタ 2');
          expect(shapes[0].visible).to.equal(false);
        });
    });

    it('assigns ids and toggles visibility via worksheet helpers', () => {
      const wb = new ExcelJS.Workbook();
      let wb2;

      return wb.xlsx
        .readFile(WITH_LINE_FIXTURE)
        .then(() => {
          const ws = wb.getWorksheet('Sheet1');
          const [shape] = ws.getShapes();
          ws.setShapeId(shape, 'line-id-001');
          expect(ws.getShapeById('line-id-001')).to.equal(shape);
          ws.hideShape('line-id-001');
          expect(shape.visible).to.equal(false);
          return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => {
          const ws = wb2.getWorksheet('Sheet1');
          const shape = ws.getShapeById('line-id-001');
          expect(shape).to.not.be.undefined();
          expect(shape.visible).to.equal(false);
          ws.showShape('line-id-001');
          expect(ws.getShapeById('line-id-001').visible).to.equal(true);
        });
    });
  });
});

describe('Parsing text body', () => {
  function addAndGetShapeWithTextBody(textBody) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    ws.addShape(
      {
        textBody,
        type: 'rect',
      },
      'B2:D6'
    );
    return ws.getShapes()[0];
  }

  it('single string', () => {
    const shape = addAndGetShapeWithTextBody('foo');
    expect(shape.props.textBody).to.deep.equal({
      paragraphs: [{runs: [{text: 'foo'}]}],
    });
  });
  it('array of strings', () => {
    const shape = addAndGetShapeWithTextBody(['foo', 'bar']);
    expect(shape.props.textBody).to.deep.equal({
      paragraphs: [{runs: [{text: 'foo'}]}, {runs: [{text: 'bar'}]}],
    });
  });
  it('array of array of strings', () => {
    const shape = addAndGetShapeWithTextBody([
      ['foo', 'bar'],
      ['baz', 'qux'],
    ]);
    expect(shape.props.textBody).to.deep.equal({
      paragraphs: [
        {runs: [{text: 'foo'}, {text: 'bar'}]},
        {runs: [{text: 'baz'}, {text: 'qux'}]},
      ],
    });
  });
  it('object', () => {
    const obj = {
      paragraphs: [
        {
          runs: [
            {
              text: 'foo',
              font: {size: 15, bold: true, italic: true, underline: 'sng'},
            },
            {
              text: 'bar',
              font: {color: {theme: 'accent1'}},
            },
          ],
          alignment: 'ctr',
        },
        {runs: [{text: 'baz'}, {text: 'qux'}]},
      ],
      vertAlign: 'b',
    };
    const shape = addAndGetShapeWithTextBody(obj);
    expect(shape.props.textBody).to.deep.equal(obj);
  });
});
