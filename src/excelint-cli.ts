// Process Excel files (input from .xls or .xlsx) with ExceLint.
// by Emery Berger, Microsoft Research / University of Massachusetts Amherst
// www.emeryberger.com

'use strict';
const fs = require('fs');
const path = require('path');
import {ExcelJSON} from './exceljson';
import {ExcelUtils} from './excelutils';
import {Colorize} from './colorize';
import {Timer} from './timer';
import {string} from 'prop-types';

const usageString = 'Usage: $0 <command> [options]';
const defaultFormattingDiscount = Colorize.getFormattingDiscount();
const defaultReportingThreshold = Colorize.getReportingThreshold();
const defaultMaxCategories = Colorize.maxCategories; // FIXME should be an accessor
const defaultMinFixSize = Colorize.minFixSize;
const defaultMaxEntropy = Colorize.maxEntropy;

let numWorkbooks = 0;
let numWorkbooksWithFormulas = 0;
let numWorkbooksWithErrors = 0;

let numSheets = 0;
let numSheetsWithErrors = 0;

// Process command-line arguments.
const args = require('yargs')
  .usage(usageString)
  .command('input', 'Input from FILENAME (.xls / .xlsx file).')
  .alias('i', 'input')
  .nargs('input', 1)
  .command(
    'directory',
    'Read from a directory of files (all ending in .xls / .xlsx).'
  )
  .alias('d', 'directory')
  .command(
    'formattingDiscount',
    'Set discount for formatting differences (default = ' +
      defaultFormattingDiscount +
      ').'
  )
  .command(
    'reportingThreshold',
    'Set the threshold % for reporting anomalous formulas (default = ' +
      defaultReportingThreshold +
      ').'
  )
  .command('suppressOutput', "Don't output the processed JSON to stdout.")
  .command(
    'noElapsedTime',
    'Suppress elapsed time output (for regression testing).'
  )
  .command(
    'maxCategories',
    'Maximum number of categories for reported errors (default = ' +
      defaultMaxCategories +
      ').'
  )
  .command(
    'minFixSize',
    'Minimum size of a fix in number of cells (default = ' +
      defaultMinFixSize +
      ')'
  )
  .command(
    'maxEntropy',
    'Maximum entropy of a proposed fix (default = ' + defaultMaxEntropy + ')'
  )
  .command('suppressFatFix', '')
  .command('suppressDifferentReferentCount', '')
  .command('suppressRecurrentFormula', '')
  .command('suppressOneExtraConstant', '')
  .command('suppressNumberOfConstantsMismatch', '')
  .command('suppressBothConstants', '')
  .command('suppressOneIsAllConstants', '')
  .command('suppressR1C1Mismatch', '')
  .command('suppressAbsoluteRefMismatch', '')
  .command('suppressOffAxisReference', '')
  .command(
    'sweep',
    'Perform a parameter sweep and report the best settings overall.'
  )
  .help('h')
  .alias('h', 'help').argv;

if (args.help) {
  process.exit(0);
}

let allFiles = [];

if (args.directory) {
  // Load up all files to process.
  allFiles = fs
    .readdirSync(args.directory)
    .filter((x: string) => x.endsWith('.xls') || x.endsWith('.xlsx'));
}
//console.log(JSON.stringify(allFiles));

// argument:
// input = filename. Default file is standard input.
let fname = '/dev/stdin';
if (args.input) {
  fname = args.input;
  allFiles = [fname];
}

if (!args.directory && !args.input) {
  console.warn('Must specify either --directory or --input.');
  process.exit(-1);
}

// argument:
// formattingDiscount = amount of impact of formatting on fix reporting (0-100%).
let formattingDiscount = defaultFormattingDiscount;
if ('formattingDiscount' in args) {
  formattingDiscount = args.formattingDiscount;
}
// Ensure formatting discount is within range (0-100, inclusive).
if (formattingDiscount < 0) {
  formattingDiscount = 0;
}
if (formattingDiscount > 100) {
  formattingDiscount = 100;
}
Colorize.setFormattingDiscount(formattingDiscount);

if (args.suppressFatFix) {
  Colorize.suppressFatFix = true;
}
if (args.suppressDifferentReferentCount) {
  Colorize.suppressDifferentReferentCount = true;
}
if (args.suppressRecurrentFormula) {
  Colorize.suppressRecurrentFormula = true;
}
if (args.suppressOneExtraConstant) {
  Colorize.suppressOneExtraConstant = true;
}
if (args.suppressNumberOfConstantsMismatch) {
  Colorize.suppressNumberOfConstantsMismatch = true;
}
if (args.suppressBothConstants) {
  Colorize.suppressBothConstants = true;
}
if (args.suppressOneIsAllConstants) {
  Colorize.suppressOneIsAllConstants = true;
}
if (args.suppressR1C1Mismatch) {
  Colorize.suppressR1C1Mismatch = true;
}
if (args.suppressAbsoluteRefMismatch) {
  Colorize.suppressAbsoluteRefMismatch = true;
}
if (args.suppressOffAxisReference) {
  Colorize.suppressOffAxisReference = true;
}

// As above, but for reporting threshold.
let reportingThreshold = defaultReportingThreshold;
if ('reportingThreshold' in args) {
  reportingThreshold = args.reportingThreshold;
}
// Ensure formatting discount is within range (0-100, inclusive).
if (reportingThreshold < 0) {
  reportingThreshold = 0;
}
if (reportingThreshold > 100) {
  reportingThreshold = 100;
}
Colorize.setReportingThreshold(reportingThreshold);

if ('maxCategories' in args) {
  Colorize.maxCategories = args.maxCategories;
}

if ('minFixSize' in args) {
  Colorize.minFixSize = args.minFixSize;
}

let maxEntropy = defaultMaxEntropy;
if ('maxEntropy' in args) {
  maxEntropy = args.maxEntropy;
  // Entropy must be between 0 and 1.
  if (maxEntropy < 0.0) {
    maxEntropy = 0.0;
  }
  if (maxEntropy > 1.0) {
    maxEntropy = 1.0;
  }
}

//
// Ready to start processing.
//

let inp = null;

let annotated_bugs = '{}';
try {
  annotated_bugs = fs.readFileSync('annotations-processed.json');
} catch (e) {}

const theBugs = JSON.parse(annotated_bugs);

let base = '';
if (args.directory) {
  base = args.directory + '/';
}

let parameters = [];
if (args.sweep) {
  const step = 10;
  for (let i = 0; i <= 100; i += step) {
    for (let j = 0; j <= 100; j += step) {
      parameters.push([i, j]);
    }
  }
} else {
  parameters = [[formattingDiscount, reportingThreshold]];
}

const f1scores = [];
const outputs = [];

for (const parms of parameters) {
  formattingDiscount = parms[0];
  Colorize.setFormattingDiscount(formattingDiscount);
  reportingThreshold = parms[1];
  Colorize.setReportingThreshold(reportingThreshold);

  const scores = [];

  for (const fname of allFiles) {
    numWorkbooks += 1;
    // Read from file.
    console.warn('processing ' + fname);
    inp = ExcelJSON.processWorkbook(base, fname);

    {
      let hasError = false;
      let hasFormula = false;
      for (let i = 0; i < inp.worksheets.length; i++) {
        const sheet = inp.worksheets[i];
        numSheets += 1;
        const workbookBasename = path.basename(inp['workbookName']);
        if (workbookBasename in theBugs) {
          if (sheet.sheetName in theBugs[workbookBasename]) {
            if (theBugs[workbookBasename][sheet.sheetName]['bugs'].length > 0) {
              hasError = true;
              numSheetsWithErrors += 1;
            }
          }
        }
        if (sheet.formulas.length > 2) {
          // ExceLint can't ever report an error if there are fewer than 3 formulas.
          hasFormula = true;
        }
      }
      if (hasError) {
        numWorkbooksWithErrors += 1;
      }
      if (hasFormula) {
        numWorkbooksWithFormulas += 1;
      }
    }

    const output = Colorize.process_workbook(inp, ""); // no bug processing for now; get all sheets
    outputs.push(output);
  }
  let averageScores = 0;
  let sumScores = 0;
  if (scores.length > 0) {
    averageScores = scores.reduce((a, b) => a + b, 0) / scores.length;
    sumScores = scores.reduce((a, b) => a + b, 0);
  }
  f1scores.push([formattingDiscount, reportingThreshold, sumScores]);
}
f1scores.sort((a, b) => {
  if (a[2] < b[2]) {
    return -1;
  }
  if (a[2] > b[2]) {
    return 1;
  }
  return 0;
});

if (!args.suppressOutput) {
  console.log(JSON.stringify(outputs, null, '\t'));
}
