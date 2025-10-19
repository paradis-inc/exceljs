const Workbook = require('../lib/doc/workbook');

const filename = process.argv[2];

if (!filename) {
  console.error('Usage: node testRowBreaks.js <excel-file-path>');
  process.exit(1);
}

const workbook = new Workbook();
workbook.xlsx
  .readFile(filename)
  .then(() => {
    console.log(`File: ${filename}\n`);

    workbook.eachSheet(worksheet => {
      console.log(`\n${'='.repeat(60)}`);
      console.log(`Sheet ${worksheet.id}: "${worksheet.name}"`);
      console.log(`Dimensions: ${JSON.stringify(worksheet.dimensions)}`);
      console.log(`${'='.repeat(60)}`);

      // Check for row breaks
      if (worksheet.rowBreaks && worksheet.rowBreaks.length > 0) {
        console.log(`\n✓ Row breaks found: ${worksheet.rowBreaks.length}`);
        console.log('Row break positions:');
        worksheet.rowBreaks.forEach((breakInfo, index) => {
          const rowNum = typeof breakInfo === 'object' ? (breakInfo.id || breakInfo.row || JSON.stringify(breakInfo)) : breakInfo;
          console.log(`  [${index + 1}] Row: ${rowNum}`);
        });
      } else {
        console.log('\n✗ No row breaks found');
      }

      // Check for column breaks as well
      if (worksheet.colBreaks && worksheet.colBreaks.length > 0) {
        console.log(`\n✓ Column breaks found: ${worksheet.colBreaks.length}`);
        console.log('Column break positions:');
        worksheet.colBreaks.forEach((breakInfo, index) => {
          console.log(`  [${index + 1}] Column: ${breakInfo.id || breakInfo}`);
        });
      } else {
        console.log('\n✗ No column breaks found');
      }

      // Print full breaks object for debugging
      console.log('\nRaw breaks data:');
      console.log('  rowBreaks:', JSON.stringify(worksheet.rowBreaks, null, 2));
      console.log('  colBreaks:', JSON.stringify(worksheet.colBreaks, null, 2));
    });

    console.log(`\n${'='.repeat(60)}\n`);
  })
  .catch(error => {
    console.error('Error reading file:');
    console.error(error.message);
    console.error(error.stack);
    process.exit(1);
  });
