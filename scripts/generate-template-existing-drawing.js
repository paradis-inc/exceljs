const path = require('path');

const ExcelJS = require('../excel');

const IMAGE_FILENAME = path.resolve(__dirname, '../spec/integration/data/image.png');
const OUTPUT_FILENAME = path.resolve(__dirname, '../spec/integration/data/test-template-existing-drawing.xlsx');

async function main() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Template');

  worksheet.getCell('A1').value = 'existing drawing template';

  const imageId = workbook.addImage({
    filename: IMAGE_FILENAME,
    extension: 'png',
  });

  worksheet.addImage(imageId, 'B2:D6');

  await workbook.xlsx.writeFile(OUTPUT_FILENAME);
  process.stdout.write(`${OUTPUT_FILENAME}\n`);
}

main().catch(error => {
  process.stderr.write(`${error.stack}\n`);
  process.exitCode = 1;
});
