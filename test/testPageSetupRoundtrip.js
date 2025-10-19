const Workbook = require('../lib/doc/workbook');
const fs = require('fs');
const path = require('path');

const filename = process.argv[2];
const outputFilename = process.argv[3];

if (!filename) {
  console.error('Usage: node testPageSetupRoundtrip.js <input-excel-file> [output-excel-file]');
  console.error('  If output file is not specified, it will be saved as <input>_roundtrip.xlsx');
  process.exit(1);
}

// Generate output filename if not provided
const actualOutputFilename = outputFilename || (() => {
  const dir = path.dirname(filename);
  const ext = path.extname(filename);
  const base = path.basename(filename, ext);
  return path.join(dir, `${base}_roundtrip${ext}`);
})();

async function testRoundtrip() {
  try {
    console.log('Step 1: Reading original file...');
    const workbook1 = new Workbook();
    await workbook1.xlsx.readFile(filename);

    const originalPageSetup = {};
    workbook1.eachSheet((worksheet, id) => {
      originalPageSetup[id] = {
        name: worksheet.name,
        pageSetup: worksheet.pageSetup ? JSON.parse(JSON.stringify(worksheet.pageSetup)) : undefined,
      };
      console.log(`\nOriginal Sheet ${id}: "${worksheet.name}"`);
      if (worksheet.pageSetup) {
        console.log('  Page Setup:');
        console.log(`    ${JSON.stringify(worksheet.pageSetup, null, 2).split('\n').join('\n    ')}`);
      } else {
        console.log('  No page setup');
      }
    });

    console.log('\n' + '='.repeat(60));
    console.log('Step 2: Writing to buffer...');
    const buffer = await workbook1.xlsx.writeBuffer();
    console.log(`Buffer size: ${buffer.length} bytes`);

    console.log('\n' + '='.repeat(60));
    console.log('Step 2b: Saving to file...');
    fs.writeFileSync(actualOutputFilename, buffer);
    console.log(`Saved to: ${actualOutputFilename}`);

    console.log('\n' + '='.repeat(60));
    console.log('Step 3: Reading from buffer...');
    const workbook2 = new Workbook();
    await workbook2.xlsx.readFile(actualOutputFilename);

    let allMatch = true;
    workbook2.eachSheet((worksheet, id) => {
      console.log(`\nRe-read Sheet ${id}: "${worksheet.name}"`);

      const original = originalPageSetup[id];
      if (!original) {
        console.log('  ⚠️  Warning: Sheet not found in original');
        allMatch = false;
        return;
      }

      // Check page setup
      const hasPageSetup = worksheet.pageSetup !== undefined && worksheet.pageSetup !== null;
      const hadPageSetup = original.pageSetup !== undefined && original.pageSetup !== null;

      if (hasPageSetup) {
        console.log('  Page Setup:');
        console.log(`    ${JSON.stringify(worksheet.pageSetup, null, 2).split('\n').join('\n    ')}`);
      } else {
        console.log('  No page setup');
      }

      // Compare
      if (hadPageSetup !== hasPageSetup) {
        console.log('  ❌ Page setup existence mismatch!');
        allMatch = false;
      } else if (hadPageSetup) {
        const origStr = JSON.stringify(original.pageSetup);
        const currStr = JSON.stringify(worksheet.pageSetup);

        if (origStr !== currStr) {
          console.log('  ❌ Page setup mismatch!');
          console.log('     Original:');
          console.log(`       ${JSON.stringify(original.pageSetup, null, 2).split('\n').join('\n       ')}`);
          console.log('     Re-read:');
          console.log(`       ${JSON.stringify(worksheet.pageSetup, null, 2).split('\n').join('\n       ')}`);
          allMatch = false;
        } else {
          console.log('  ✅ Page setup matches perfectly!');
        }
      } else {
        console.log('  ✅ No page setup in both (match)');
      }
    });

    console.log('\n' + '='.repeat(60));
    console.log(`Output file: ${actualOutputFilename}`);
    console.log('='.repeat(60));
    if (allMatch) {
      console.log('✅ SUCCESS: All page setup preserved correctly!');
      console.log(`\nYou can open ${actualOutputFilename} to verify the page setup is preserved.`);
      process.exit(0);
    } else {
      console.log('❌ FAILURE: Some page setup was not preserved!');
      console.log(`\nCheck ${actualOutputFilename} to see what was written.`);
      process.exit(1);
    }
  } catch (error) {
    console.error('Error during roundtrip test:');
    console.error(error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

testRoundtrip();
