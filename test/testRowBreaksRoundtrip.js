const Workbook = require('../lib/doc/workbook');
const fs = require('fs');
const path = require('path');

const filename = process.argv[2];
const outputFilename = process.argv[3];

if (!filename) {
  console.error('Usage: node testRowBreaksRoundtrip.js <input-excel-file> [output-excel-file]');
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

    const originalBreaks = {};
    workbook1.eachSheet((worksheet, id) => {
      originalBreaks[id] = {
        name: worksheet.name,
        rowBreaks: worksheet.rowBreaks ? JSON.parse(JSON.stringify(worksheet.rowBreaks)) : undefined,
        colBreaks: worksheet.colBreaks ? JSON.parse(JSON.stringify(worksheet.colBreaks)) : undefined,
      };
      console.log(`\nOriginal Sheet ${id}: "${worksheet.name}"`);
      if (worksheet.rowBreaks && worksheet.rowBreaks.length > 0) {
        console.log(`  Row breaks: ${worksheet.rowBreaks.length}`);
        worksheet.rowBreaks.forEach((brk, idx) => {
          console.log(`    [${idx + 1}] Row ${brk.id}, max: ${brk.max}, manual: ${brk.man}`);
        });
      } else {
        console.log('  No row breaks');
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

      const original = originalBreaks[id];
      if (!original) {
        console.log('  ⚠️  Warning: Sheet not found in original');
        allMatch = false;
        return;
      }

      // Check row breaks
      const hasRowBreaks = worksheet.rowBreaks && worksheet.rowBreaks.length > 0;
      const hadRowBreaks = original.rowBreaks && original.rowBreaks.length > 0;

      if (hasRowBreaks) {
        console.log(`  Row breaks: ${worksheet.rowBreaks.length}`);
        worksheet.rowBreaks.forEach((brk, idx) => {
          console.log(`    [${idx + 1}] Row ${brk.id}, max: ${brk.max}, manual: ${brk.man}`);
        });
      } else {
        console.log('  No row breaks');
      }

      // Compare
      if (hadRowBreaks !== hasRowBreaks) {
        console.log('  ❌ Row breaks existence mismatch!');
        allMatch = false;
      } else if (hadRowBreaks) {
        if (original.rowBreaks.length !== worksheet.rowBreaks.length) {
          console.log(`  ❌ Row breaks count mismatch! Original: ${original.rowBreaks.length}, Re-read: ${worksheet.rowBreaks.length}`);
          allMatch = false;
        } else {
          let breaksMatch = true;
          for (let i = 0; i < original.rowBreaks.length; i++) {
            const orig = original.rowBreaks[i];
            const curr = worksheet.rowBreaks[i];
            if (orig.id !== curr.id || orig.max !== curr.max || orig.man !== curr.man) {
              console.log(`  ❌ Row break [${i}] mismatch!`);
              console.log(`     Original: ${JSON.stringify(orig)}`);
              console.log(`     Re-read:  ${JSON.stringify(curr)}`);
              breaksMatch = false;
              allMatch = false;
            }
          }
          if (breaksMatch) {
            console.log('  ✅ Row breaks match perfectly!');
          }
        }
      } else {
        console.log('  ✅ No row breaks in both (match)');
      }
    });

    console.log('\n' + '='.repeat(60));
    console.log(`Output file: ${actualOutputFilename}`);
    console.log('='.repeat(60));
    if (allMatch) {
      console.log('✅ SUCCESS: All row breaks preserved correctly!');
      console.log(`\nYou can open ${actualOutputFilename} to verify the row breaks are preserved.`);
      process.exit(0);
    } else {
      console.log('❌ FAILURE: Some row breaks were not preserved!');
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
