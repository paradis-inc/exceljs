#!/usr/bin/env node
/*
 * CLI helper to assign IDs to shapes and toggle their visibility.
 *
 * Examples:
 *   node scripts/shape-labeler.js --input test/data/withline.xlsx --sheet Sheet1 --index 0 --set-id line-01 --hide
 *   node scripts/shape-labeler.js --input test/data/withline.xlsx --sheet Sheet1 --id line-01 --show --output test/data/withline.visible.xlsx
 *   node scripts/shape-labeler.js --input test/data/withline.xlsx --sheet Sheet1 --list
 */

const path = require('path');
const ExcelJS = require('../excel.js');

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      // treat as positional: input first, sheet second
      if (!args.input) {
        args.input = arg;
      } else if (!args.sheet) {
        args.sheet = arg;
      }
      continue;
    }
    const key = arg.slice(2);
    const next = argv[i + 1];
    switch (key) {
      case 'input':
        args.input = next;
        i += 1;
        break;
      case 'output':
        args.output = next;
        i += 1;
        break;
      case 'sheet':
        args.sheet = next;
        i += 1;
        break;
      case 'id':
        args.id = next;
        i += 1;
        break;
      case 'set-id':
        args.setId = next;
        i += 1;
        break;
      case 'index':
        args.index = parseInt(next, 10);
        i += 1;
        break;
      case 'visible':
        args.visible = next === 'true';
        i += 1;
        break;
      case 'hide':
        args.hide = true;
        break;
      case 'show':
        args.show = true;
        break;
      case 'list':
        args.list = true;
        break;
      default:
        throw new Error(`Unknown option: --${key}`);
    }
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.input) {
    throw new Error('Please provide --input <file>');
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(args.input);
  const ws = args.sheet ? wb.getWorksheet(args.sheet) : wb.worksheets[0];
  if (!ws) {
    throw new Error('Worksheet not found');
  }

  const shapes = ws.getShapes();
  if (!shapes.length) {
    console.log('No shapes found');
    return;
  }

  if (args.list) {
    shapes.forEach((shape, idx) => {
      console.log(`#${idx}: name=${shape.name || '(none)'} visible=${shape.visible !== false}`);
    });
    return;
  }

  let shape;
  if (args.id) {
    shape = ws.getShapeById(args.id);
    if (!shape) {
      throw new Error(`Shape with id "${args.id}" not found`);
    }
  } else if (typeof args.index === 'number') {
    shape = shapes[args.index];
  } else {
    shape = shapes[0];
  }

  if (!shape) {
    throw new Error('Target shape not found');
  }

  if (args.setId) {
    ws.setShapeId(shape, args.setId);
  }

  if (args.hide) {
    ws.hideShape(shape);
  } else if (args.show) {
    ws.showShape(shape);
  } else if (typeof args.visible === 'boolean') {
    ws.setShapeVisibility(shape, args.visible);
  }

  const output = args.output || args.input;
  await wb.xlsx.writeFile(output);
  console.log(`Updated shape saved to ${path.resolve(output)}`);
}

main().catch(error => {
  console.error(error.message);
  process.exit(1);
});
