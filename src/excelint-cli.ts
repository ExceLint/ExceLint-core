// Process Excel files (input from .xls or .xlsx) with ExceLint.
// by Emery Berger, Microsoft Research / University of Massachusetts Amherst
// www.emeryberger.com

'use strict';
let fs = require('fs');
let path = require('path');
import { ExcelJSON } from './exceljson';
import { ExcelUtils } from './excelutils';
import { Colorize } from './colorize';
import { Timer } from './timer';
import { string } from 'prop-types';

enum BinCategories {
    RecurrentFormula = "recurrent-formula", // formulas refer to each other
    OneExtraConstant = "one-extra-constant", // one has no constant and the other has one constant
    NumberOfConstantsMismatch = "number-of-constants-mismatch", // both have constants but not the same number of constants
    BothConstants = "both-constants", // both have only constants but differ in numeric value
    OneIsAllConstants = "one-is-all-constants", // one is entirely constants and other is formula
    AbsoluteRefMismatch = "absolute-ref-mismatch", // relative vs. absolute mismatch
    OffAxisReference = "off-axis-reference", // references refer to different columns or rows
    R1C1Mismatch = "r1c1-mismatch", // different R1C1 representations
    DifferentReferentCount = "different-referent-count", // ranges have different number of referents
    // Not yet implemented.
    RefersToEmptyCells = "refers-to-empty-cells",
    UsesDifferentOperations = "uses-different-operations", // e.g. SUM vs. AVERAGE
    // Fall-through category
    Unclassified = "unclassified",
}

type excelintVector = [number, number, number];

// Convert a rectangle into a list of indices.
function expand(first: excelintVector, second: excelintVector): Array<excelintVector> {
    const [fcol, frow] = first;
    const [scol, srow] = second;
    let expanded: Array<excelintVector> = [];
    for (let i = fcol; i <= scol; i++) {
        for (let j = frow; j <= srow; j++) {
            expanded.push([i, j, 0]);
        }
    }
    return expanded;
}

// Set to true to use the hard-coded example below.
const useExample = false;

const usageString = 'Usage: $0 <command> [options]';
const defaultFormattingDiscount = Colorize.getFormattingDiscount();
const defaultReportingThreshold = Colorize.getReportingThreshold();
const defaultMaxCategories = 2;

let numWorkbooks = 0;
let numWorkbooksWithFormulas = 0;
let numWorkbooksWithErrors = 0;
let numSheets = 0;
let numSheetsWithErrors = 0;
let sheetTruePositives = 0;
let sheetFalsePositives = 0;

// Process command-line arguments.
const args = require('yargs')
    .usage(usageString)
    .command('input', 'Input from FILENAME (.xls / .xlsx file).')
    .alias('i', 'input')
    .nargs('input', 1)
    .command('directory', 'Read from a directory of files (all ending in .xls / .xlsx).')
    .alias('d', 'directory')
    .command('formattingDiscount', 'Set discount for formatting differences (default = ' + defaultFormattingDiscount + ').')
    .command('reportingThreshold', 'Set the threshold % for reporting anomalous formulas (default = ' + defaultReportingThreshold + ').')
    .command('suppressOutput', 'Don\'t output the processed JSON to stdout.')
    .command('noElapsedTime', 'Suppress elapsed time output (for regression testing).')
    .command('maxCategories', 'Maximum number of categories for reported errors (default = ' + defaultMaxCategories + ').')
    .command('suppressDifferentReferentCount', '')
    .command('suppressRecurrentFormula', '')
    .command('suppressOneExtraConstant', '')
    .command('suppressNumberOfConstantsMismatch', '')
    .command('suppressBothConstants', '')
    .command('suppressOneIsAllConstants', '')
    .command('suppressR1C1Mismatch', '')
    .command('suppressAbsoluteRefMismatch', '')
    .command('suppressOffAxisReference', '')
    .command('sweep', 'Perform a parameter sweep and report the best settings overall.')
    .help('h')
    .alias('h', 'help')
    .argv;

if (args.help) {
    process.exit(0);
}

let allFiles = [];

if (args.directory) {
    // Load up all files to process.
    allFiles = fs.readdirSync(args.directory).filter((x: string) => x.endsWith('.xls') || x.endsWith('.xlsx'));
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

if (!('maxCategories' in args)) {
  args.maxCategories = defaultMaxCategories;
}

//
// Ready to start processing.
//

let inp = null;

if (useExample) {
    // A simple example.
    inp = {
        workbookName: 'example',
        worksheets: [{
            sheetname: 'Sheet1',
            usedRangeAddress: 'Sheet1!E12:E21',
            formulas: [
                ['=D12'], ['=D13'],
                ['=D14'], ['=D15'],
                ['=D16'], ['=D17'],
                ['=D18'], ['=D19'],
                ['=D20'], ['=C21']
            ],
            values: [
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0'],
                ['0'], ['0']
            ],
            styles: [
                [''], [''],
                [''], [''],
                [''], [''],
                [''], [''],
                [''], ['']
            ]
        }]
    };
}

let annotated_bugs = '{}';
try {
    annotated_bugs = fs.readFileSync('annotations-processed.json');
} catch (e) {
}

let bugs = JSON.parse(annotated_bugs);

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

let f1scores = [];
let outputs = [];

for (let parms of parameters) {
    formattingDiscount = parms[0];
    Colorize.setFormattingDiscount(formattingDiscount);
    reportingThreshold = parms[1];
    Colorize.setReportingThreshold(reportingThreshold);

    let scores = [];

    for (let fname of allFiles) {
	numWorkbooks += 1;
        // Read from file.
        console.warn('processing ' + fname);
        inp = ExcelJSON.processWorkbook(base, fname);

        let output = {
            'workbookName': path.basename(inp['workbookName']),
            'worksheets': {}
        };

	{
	    let hasError = false;
	    let hasFormula = false;
            for (let i = 0; i < inp.worksheets.length; i++) {
		const sheet = inp.worksheets[i];
		numSheets += 1;
		const workbookBasename = path.basename(inp['workbookName']);
		if (workbookBasename in bugs) {
		    if (sheet.sheetName in bugs[workbookBasename]) {
			hasError = true;
			numSheetsWithErrors += 1;
		    }
		}
		if (sheet.formulas.length > 2) { // ExceLint can't ever report an error if there are fewer than 3 formulas.
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
	
        for (let i = 0; i < inp.worksheets.length; i++) {
            const sheet = inp.worksheets[i];
	    
            // Skip empty sheets.
            if ((sheet.formulas.length === 0) && (sheet.values.length === 0)) {
                continue;
            }
	    console.warn(output['workbookName'] + " - " + sheet.sheetName);
	    
            // Get rid of multiple exclamation points in the used range address,
            // as these interfere with later regexp parsing.
            let usedRangeAddress = sheet.usedRangeAddress;
            usedRangeAddress = usedRangeAddress.replace(/!(!+)/, '!');

            const myTimer = new Timer('excelint');

            // Get anomalous cells and proposed fixes, among others.
            let [anomalous_cells, grouped_formulas, grouped_data, proposed_fixes]
                = Colorize.process_suspicious(usedRangeAddress, sheet.formulas, sheet.values);

            // Adjust the fixes based on font stuff. We should allow parameterization here for weighting (as for thresholding).
            // NB: origin_col and origin_row currently hard-coded at 0,0.

            proposed_fixes = Colorize.adjust_proposed_fixes(proposed_fixes, sheet.styles, 0, 0);

            // Adjust the proposed fixes for real (just adjusting the scores downwards by the formatting discount).
            let initial_adjusted_fixes = [];
	    let final_adjusted_fixes = []; // We will eventually trim these.
            // tslint:disable-next-line: forin
            for (let ind = 0; ind < proposed_fixes.length; ind++) {
                const f = proposed_fixes[ind];
                const [score, first, second, sameFormat] = f;
                let adjusted_score = -score;
                if (!sameFormat) {
                    adjusted_score *= (100 - formattingDiscount) / 100;
                }
                if (adjusted_score * 100 >= reportingThreshold) {
                    initial_adjusted_fixes.push([adjusted_score, first, second]);
                }
            }

	    // Process all the fixes, classifying and optionally pruning them.

            let example_fixes_r1c1 = [];
            for (let ind = 0; ind < initial_adjusted_fixes.length; ind++) {
		// Determine the direction of the range (vertical or horizontal) by looking at the axes.
                let direction = "";
                if (initial_adjusted_fixes[ind][1][0][0] === initial_adjusted_fixes[ind][2][0][0]) {
                    direction = "vertical";
                } else {
                    direction = "horizontal";
                }
                let formulas = [];              // actual formulas
                let print_formulas = [];        // formulas with a preface (the cell name containing each)
                let r1c1_formulas = [];         // formulas in R1C1 format
                let r1c1_print_formulas = [];   // as above, but for R1C1 formulas
                let all_numbers = [];           // all the numeric constants in each formula
                let numbers = [];               // the sum of all the numeric constants in each formula
                let dependence_count = [];      // the number of dependent cells
                let absolute_refs = [];         // the number of absolute references in each formula
                let dependence_vectors = [];
		// Generate info about the formulas.
                for (let i = 0; i < 2; i++) {
                    // the coordinates of the cell containing the first formula in the proposed fix range
                    const formulaCoord = initial_adjusted_fixes[ind][i + 1][0];
                    const formulaX = formulaCoord[1] - 1;                   // row
                    const formulaY = formulaCoord[0] - 1;                   // column
                    const formula = sheet.formulas[formulaX][formulaY];   // the formula itself
                    const numeric_constants = ExcelUtils.numeric_constants(formula); // all numeric constants in the formula
                    all_numbers.push(numeric_constants);
                    numbers.push(numbers.reduce((a, b) => a + b, 0));      // the sum of all numeric constants
                    const dependences_wo_constants = ExcelUtils.all_cell_dependencies(formula, formulaY + 1, formulaX + 1, false);
                    dependence_count.push(dependences_wo_constants.length);
                    const r1c1 = ExcelUtils.formulaToR1C1(formula, formulaY + 1, formulaX + 1);
                    const preface = ExcelUtils.column_index_to_name(formulaY + 1) + (formulaX + 1) + ":";
                    const cellPlusFormula = preface + r1c1;
                    // Add the formulas plus their prefaces (the latter for printing).
                    r1c1_formulas.push(r1c1);
                    r1c1_print_formulas.push(cellPlusFormula);
                    formulas.push(formula);
                    print_formulas.push(preface + formula);
                    absolute_refs.push((formula.match(/\$/g) || []).length);
                    // console.log(preface + JSON.stringify(dependences_wo_constants));
                    dependence_vectors.push(dependences_wo_constants);
                }
                let totalNumericDiff = Math.abs(numbers[0] - numbers[1]);
                // Binning.
                let bin = [];
                // Check for recurrent formulas.
                for (let i = 0; i < dependence_vectors.length; i++) {
                    // If there are at least two dependencies and one of them is -1 in the column (row),
                    // it is a recurrence (the common recurrence relation of starting at a value and
                    // referencing, say, =B10+1).
                    if (dependence_vectors[i].length > 0) {
                        if ((direction === "vertical") && ((dependence_vectors[i][0][0] === 0) && (dependence_vectors[i][0][1] === -1))) {
                            bin.push(BinCategories.RecurrentFormula);
                            break;
                        }
                        if ((direction === "horizontal") && ((dependence_vectors[i][0][0] === -1) && (dependence_vectors[i][0][1] === 0))) {
                            bin.push(BinCategories.RecurrentFormula);
                            break;
                        }
                    }
                }
                // Different number of referents (dependencies).
                if (dependence_count[0] !== dependence_count[1]) {
                    bin.push(BinCategories.DifferentReferentCount);
                }
                // Different number of constants.
                if (all_numbers[0].length !== all_numbers[1].length) {
                    if (Math.abs(all_numbers[0].length - all_numbers[1].length) === 1) {
                        bin.push(BinCategories.OneExtraConstant);
                    } else {
                        bin.push(BinCategories.NumberOfConstantsMismatch);
                    }
                }
                // Both constants.
                if ((all_numbers[0].length > 0) && (all_numbers[1].length > 0)) {
                    // Both have numbers.
                    if (dependence_count[0] + dependence_count[1] === 0) {
                        // Both have no dependents.
                        bin.push(BinCategories.BothConstants);
                    } else {
                        if (dependence_count[0] * dependence_count[1] === 0) {
                            // One is a constant.
                            bin.push(BinCategories.OneIsAllConstants);
                        }
                    }
                }
                // Mismatched R1C1 representation.
                if (r1c1_formulas[0] !== r1c1_formulas[1]) {
                    // The formulas don't match, but it could
                    // be because of the presence of (possibly
                    // different) constants instead of the
                    // dependencies being different. Do a deep comparison
                    // here.
                    if (JSON.stringify(dependence_vectors[0].sort()) !== JSON.stringify(dependence_vectors[1].sort())) {
                        bin.push(BinCategories.R1C1Mismatch);
                    }
                }
                // Different number of absolute ($, a.k.a. "anchor") references.
                if (absolute_refs[0] !== absolute_refs[1]) {
                    bin.push(BinCategories.AbsoluteRefMismatch);
                }
                // Dependencies that are neither vertical or horizontal (likely errors if an absolute-ref-mismatch).
                for (let i = 0; i < dependence_vectors.length; i++) {
                    if (dependence_vectors[i].length > 0) {
                        if (dependence_vectors[i][0][0] * dependence_vectors[i][0][1] !== 0) {
                            bin.push(BinCategories.OffAxisReference);
                            break;
                        }
                    }
                }
                if (bin === []) {
                    bin.push(BinCategories.Unclassified);
                }
		// IMPORTANT:
		// Exclude reported bugs subject to certain conditions.
		if ((bin.length > args.maxCategories) // Too many categories
		    || ((bin.indexOf(BinCategories.DifferentReferentCount) != -1) && args.suppressDifferentReferentCount)
		    || ((bin.indexOf(BinCategories.RecurrentFormula) != -1) && args.suppressRecurrentFormula)
		    || ((bin.indexOf(BinCategories.OneExtraConstant) != -1) && args.suppressOneExtraConstant)
		    || ((bin.indexOf(BinCategories.NumberOfConstantsMismatch) != -1) && args.suppressNumberOfConstantsMismatch)
		    || ((bin.indexOf(BinCategories.BothConstants) != -1) && args.suppressBothConstants)
		    || ((bin.indexOf(BinCategories.OneIsAllConstants) != -1) && args.suppressOneIsAllConstants)
		    || ((bin.indexOf(BinCategories.R1C1Mismatch) != -1) && args.suppressR1C1Mismatch)
		    || ((bin.indexOf(BinCategories.AbsoluteRefMismatch) != -1) && args.suppressAbsoluteRefMismatch)
		    || ((bin.indexOf(BinCategories.OffAxisReference) != -1) && args.suppressOffAxisReference))
		{
		    console.warn("Omitted " + JSON.stringify(print_formulas) + "(" + JSON.stringify(bin) + ")");
		    continue;
		} else {
		    console.warn("NOT omitted " + JSON.stringify(print_formulas) + "(" + JSON.stringify(bin) + ")");
		}
		final_adjusted_fixes.push(initial_adjusted_fixes[ind]);
		
                example_fixes_r1c1.push({
                    "bin": bin,
                    "direction": direction,
                    "numbers": numbers,
                    "numeric_difference": totalNumericDiff,
                    "magnitude_numeric_difference": (totalNumericDiff === 0) ? 0 : Math.log10(totalNumericDiff),
                    "formulas": print_formulas,
                    "r1c1formulas": r1c1_print_formulas,
                    "dependence_vectors": dependence_vectors
                });
                // example_fixes_r1c1.push([direction, formulas]);
            }

            let elapsed = myTimer.elapsedTime();
            if (args.noElapsedTime) {
                elapsed = 0; // Dummy value, used for regression testing.
            }
            // Compute number of cells containing formulas.
            const numFormulaCells = (sheet.formulas.flat().filter(x => x.length > 0)).length;

            // Count the number of non-empty cells.
            const numValueCells = (sheet.values.flat().filter(x => x.length > 0)).length;

            // Compute total number of cells in the sheet (rows * columns).
            const columns = sheet.values[0].length;
            const rows = sheet.values.length;
            const totalCells = rows * columns;

            const out = {
                'anomalousnessThreshold': reportingThreshold,
                'formattingDiscount': formattingDiscount,
                // 'proposedFixes': final_adjusted_fixes,
                'exampleFixes': example_fixes_r1c1,
                //		'exampleFixesR1C1' : example_fixes_r1c1,
                'anomalousRanges': final_adjusted_fixes.length,
                'weightedAnomalousRanges': 0, // actually calculated below.
                'anomalousCells': 0, // actually calculated below.
                'elapsedTimeSeconds': elapsed / 1e6,
                'columns': columns,
                'rows': rows,
                'totalCells': totalCells,
                'numFormulaCells': numFormulaCells,
                'numValueCells': numValueCells
            };

            // Compute precision and recall of proposed fixes, if we have annotated ground truth.
            const workbookBasename = path.basename(inp['workbookName']);
            // Build list of bugs.
            let foundBugs: any = final_adjusted_fixes.map(x => {
                if (x[0] >= (reportingThreshold / 100)) {
                    return expand(x[1][0], x[1][1]).concat(expand(x[2][0], x[2][1]));
                } else {
                    return [];
                }
            });
            const foundBugsArray: any = Array.from(new Set(foundBugs.flat(1).map(JSON.stringify)));
            foundBugs = foundBugsArray.map(JSON.parse);
            out['anomalousCells'] = foundBugs.length;
            let weightedAnomalousRanges = final_adjusted_fixes.map(x => x[0]).reduce((x, y) => x + y, 0);
            out['weightedAnomalousRanges'] = weightedAnomalousRanges;
            if (workbookBasename in bugs) {
                if (sheet.sheetName in bugs[workbookBasename]) {
                    const trueBugs = bugs[workbookBasename][sheet.sheetName]['bugs'];
                    const totalTrueBugs = trueBugs.length;
                    const trueBugsJSON = trueBugs.map(x => JSON.stringify(x));
                    const foundBugsJSON = foundBugs.map(x => JSON.stringify(x));
                    const truePositives = trueBugsJSON.filter(value => foundBugsJSON.includes(value)).map(x => JSON.parse(x));
                    const falsePositives = foundBugsJSON.filter(value => !trueBugsJSON.includes(value)).map(x => JSON.parse(x));
                    const falseNegatives = trueBugsJSON.filter(value => !foundBugsJSON.includes(value)).map(x => JSON.parse(x));
                    let precision = 0;
                    let recall = 0;
                    out['falsePositives'] = falsePositives.length;
                    out['falseNegatives'] = falseNegatives.length;
                    out['truePositives'] = truePositives.length;
		    // sheetFalsePositives just equals 1 if there are any false positives;
		    // similarly for others.
		    out['sheetFalsePositives'] = (falsePositives.length > 0) ? 1 : 0;
		    out['sheetFalseNegatives'] = (falseNegatives.length > 0) ? 1 : 0;
		    out['sheetTruePositives'] = (truePositives.length > 0) ? 1 : 0;

		    if (truePositives.length) {
			sheetTruePositives += 1;
		    }
		    if (falsePositives.length) {
			sheetFalsePositives += 1;
		    }
		    
                    // We adopt the methodology used by the ExceLint paper (OOPSLA 18):
                    //   "When a tool flags nothing, we define precision to
                    //    be 1, since the tool makes no mistakes. When a benchmark contains no errors but the tool flags
                    //    anything, we define precision to be 0 since nothing that it flags can be a real error."

                    if (foundBugs.length === 0) {
                        out['precision'] = 1;
                    }
                    if ((truePositives.length === 0) && (foundBugs.length > 0)) {
                        out['precision'] = 0;
                    }
                    if ((truePositives.length > 0) && (foundBugs.length > 0)) {
                        precision = truePositives.length / (truePositives.length + falsePositives.length);
                        out['precision'] = precision;
                    }
                    if (falseNegatives.length + trueBugs.length > 0) {
                        recall = truePositives.length / (falseNegatives.length + truePositives.length);
                        out['recall'] = recall;
                    } else {
                        // No bugs to find means perfect recall. NOTE: this is not described in the paper.
                        out['recall'] = 1;
                    }
                    scores.push(truePositives.length - falsePositives.length);
                    if (false) {
                        if (precision + recall > 0) {
                            // F1 score: https://en.wikipedia.org/wiki/F1_score
                            const f1score = (2 * precision * recall) / (precision + recall);
                            /// const f1score = precision; //// FIXME for testing (2 * precision * recall) / (precision + recall);
                            scores.push(f1score);
                        }
                    }
                }
            }
	    out['proposedFixes'] = final_adjusted_fixes;
            output.worksheets[sheet.sheetName] = out;
        }
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
f1scores.sort((a, b) => { if (a[2] < b[2]) { return -1; } if (a[2] > b[2]) { return 1; } return 0; });
// Now find the lowest threshold with the highest F1 score.
const maxScore = f1scores.reduce((a, b) => { if (a[2] > b[2]) { return a[2]; } else { return b[2]; } });
//console.log('maxScore = ' + maxScore);
// Find the first one with the max.
const firstMax = f1scores.find(item => { return item[2] === maxScore; });
//console.log('first max = ' + firstMax);
if (!args.suppressOutput) {
    console.log(JSON.stringify(outputs, null, '\t'));
}
// console.log(JSON.stringify(f1scores));

console.log("Num workbooks = " + numWorkbooks);
console.log("Num workbooks with errors = " + numWorkbooksWithErrors);
console.log("Num workbooks with formulas = " + numWorkbooksWithFormulas);
console.log("Num sheets = " + numSheets);
console.log("Num sheets with errors = " + numSheetsWithErrors);
console.log("Sheets with ExceLint true positives = " + sheetTruePositives);
console.log("Sheets with ExceLint false positives = " + sheetFalsePositives);
