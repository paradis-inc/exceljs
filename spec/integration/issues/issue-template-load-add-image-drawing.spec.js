const fs = require('fs');
const path = require('path');
const {promisify} = require('util');

const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const fsReadFileAsync = promisify(fs.readFile);

const IMAGE_FILENAME = path.resolve(__dirname, '../data/image.png');
const TEMPLATE_FILENAME = path.resolve(__dirname, '../data/test-template-existing-drawing.xlsx');

describe('github issues', () => {
  it('appends images added after loading a template with an existing drawing', async () => {
    const workbook = new ExcelJS.Workbook();
    const templateBuffer = await fsReadFileAsync(TEMPLATE_FILENAME);

    await workbook.xlsx.load(templateBuffer);

    const worksheet = workbook.getWorksheet('Template');
    const imageId = workbook.addImage({
      filename: IMAGE_FILENAME,
      extension: 'png',
    });

    worksheet.addImage(imageId, 'B10:D14');

    const outputBuffer = await workbook.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(outputBuffer);

    expect(zip.files['xl/media/image2.png']).to.not.be.undefined();

    const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string');
    const drawingRelsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string');
    const pictureCount = (drawingXml.match(/<xdr:pic>/g) || []).length;

    expect(pictureCount).to.equal(2);
    expect(drawingRelsXml).to.contain('../media/image2.png');
  });
});
