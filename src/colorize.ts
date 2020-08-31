const path = require('path');

// Polyfill for flat (IE & Edge)
const flat = require('array.prototype.flat');
flat.shim();

//import { rgb2hex, GroupedList } from 'office-ui-fabric-react';
//import { ColorUtils } from './colorutils';
import {ExcelUtils} from './excelutils';
import {RectangleUtils} from './rectangleutils';
//import { ExcelUtilities } from '@microsoft/office-js-helpers';
import {Timer} from './timer';
import {JSONclone} from './jsonclone';
import {find_all_proposed_fixes} from './groupme';
import {Stencil, InfoGain} from './infogain';

export class Colorize {
  public static maxCategories = 2; // Maximum number of categories for reported errors
  public static minFixSize = 3; // Minimum size of a fix in number of cells
  public static maxEntropy = 1.0; // Maximum entropy of a proposed fix

  // Suppressing certain categories of errors.

  public static suppressFatFix = true;
  public static suppressDifferentReferentCount = false;
    public static suppressRecurrentFormula = false; // true;
    public static suppressOneExtraConstant = false; // true;
    public static suppressNumberOfConstantsMismatch = false; // = true;
    public static suppressBothConstants = false; // true;
    public static suppressOneIsAllConstants = false; // true;
  public static suppressR1C1Mismatch = false;
  public static suppressAbsoluteRefMismatch = false;
    public static suppressOffAxisReference = false; // true;

  public static noElapsedTime = false; // if true, don't report elapsed time
    public static reportingThreshold = 0; // 35; // Percent of anomalousness
  public static suspiciousCellsReportingThreshold = 85; //  percent of bar
  public static formattingDiscount = 50; // percent of discount: 100% means different formats = not suspicious at all

  // Limits on how many formulas or values to attempt to process.
  private static formulasThreshold = 10000;
  private static valuesThreshold = 10000;

  public static setReportingThreshold(value: number) {
    Colorize.reportingThreshold = value;
  }

  public static getReportingThreshold(): number {
    return Colorize.reportingThreshold;
  }

  public static setFormattingDiscount(value: number) {
    Colorize.formattingDiscount = value;
  }

  public static getFormattingDiscount(): number {
    return Colorize.formattingDiscount;
  }

  // Color-blind friendly color palette.
  public static palette = [
    '#ecaaae',
    '#74aff3',
    '#d8e9b2',
    '#deb1e0',
    '#9ec991',
    '#adbce9',
    '#e9c59a',
    '#71cdeb',
    '#bfbb8a',
    '#94d9df',
    '#91c7a8',
    '#b4efd3',
    '#80b6aa',
    '#9bd1c6',
  ]; // removed '#73dad1'

  // True iff this class been initialized.
  private static initialized = false;

  // The array of colors (used to hash into).
  private static color_list = [];

  // A multiplier for the hash function.
  private static Multiplier = 1; // 103037;

  // A hash string indicating no dependencies.
  private static distinguishedZeroHash = '12345';

  public static initialize() {
    if (!this.initialized) {
      // Create the color palette array.
      const arr = Colorize.palette;
      for (let i = 0; i < arr.length; i++) {
        this.color_list.push(arr[i]);
      }
      this.initialized = true;
    }
  }

  // Get the color corresponding to a hash value.
  public static get_color(hashval: number): string {
    const color = this.color_list[(hashval * 1) % this.color_list.length];
    return color;
  }

  public static process_workbook(inp: any, sheetName: string): any {
    const output = {
      workbookName: path.basename(inp['workbookName']),
      worksheets: {},
    };
    /* disabled for now:
    const scores = [];
    const sheetTruePositiveSet = new Set();
    const sheetFalsePositiveSet = new Set();
    const sheetTruePositives = 0;
    const sheetFalsePositives = 0;
    */

    for (let i = 0; i < inp.worksheets.length; i++) {
      const sheet = inp.worksheets[i];
      // If sheet name argument is not "" and doesn't match, skip it.
	if ((sheetName !== "") && (sheet.sheetName !== sheetName)) {
	    continue;
	}
      // Skip empty sheets.
      if (sheet.formulas.length === 0 && sheet.values.length === 0) {
        continue;
      }
      // console.warn(output['workbookName'] + " - " + sheet.sheetName);

      // Get rid of multiple exclamation points in the used range address,
      // as these interfere with later regexp parsing.
      let usedRangeAddress = sheet.usedRangeAddress;
      usedRangeAddress = usedRangeAddress.replace(/!(!+)/, '!');

      const myTimer = new Timer('excelint');

      // Get anomalous cells and proposed fixes, among others.
      let [
        _anomalous_cells,
        grouped_formulas,
        grouped_data,
        proposed_fixes,
      ] = Colorize.process_suspicious(
        usedRangeAddress,
        sheet.formulas,
        sheet.values
      );

      // Adjust the fixes based on font stuff. We should allow parameterization here for weighting (as for thresholding).
      // NB: origin_col and origin_row currently hard-coded at 0,0.

      proposed_fixes = Colorize.adjust_proposed_fixes(
        proposed_fixes,
        sheet.styles,
        0,
        0
      );

      // Adjust the proposed fixes for real (just adjusting the scores downwards by the formatting discount).
      const initial_adjusted_fixes = [];
      const final_adjusted_fixes = []; // We will eventually trim these.
      // tslint:disable-next-line: forin
      for (let ind = 0; ind < proposed_fixes.length; ind++) {
        const f = proposed_fixes[ind];
        const [score, first, second, sameFormat] = f;
        let adjusted_score = -score;
        if (!sameFormat) {
          adjusted_score *= (100 - Colorize.formattingDiscount) / 100;
        }
        if (adjusted_score * 100 >= Colorize.reportingThreshold) {
          initial_adjusted_fixes.push([adjusted_score, first, second]);
        }
      }

      // Process all the fixes, classifying and optionally pruning them.

      const example_fixes_r1c1 = [];
      for (let ind = 0; ind < initial_adjusted_fixes.length; ind++) {
        // Determine the direction of the range (vertical or horizontal) by looking at the axes.
        let direction = '';
        if (
          initial_adjusted_fixes[ind][1][0][0] ===
          initial_adjusted_fixes[ind][2][0][0]
        ) {
          direction = 'vertical';
        } else {
          direction = 'horizontal';
        }
        const formulas = []; // actual formulas
        const print_formulas = []; // formulas with a preface (the cell name containing each)
        const r1c1_formulas = []; // formulas in R1C1 format
        const r1c1_print_formulas = []; // as above, but for R1C1 formulas
        const all_numbers = []; // all the numeric constants in each formula
        const numbers = []; // the sum of all the numeric constants in each formula
        const dependence_count = []; // the number of dependent cells
        const absolute_refs = []; // the number of absolute references in each formula
        const dependence_vectors = [];
        // Generate info about the formulas.
        for (let i = 0; i < 2; i++) {
          // the coordinates of the cell containing the first formula in the proposed fix range
          const formulaCoord = initial_adjusted_fixes[ind][i + 1][0];
          const formulaX = formulaCoord[1] - 1; // row
          const formulaY = formulaCoord[0] - 1; // column
          const formula = sheet.formulas[formulaX][formulaY]; // the formula itself
          const numeric_constants = ExcelUtils.numeric_constants(formula); // all numeric constants in the formula
          all_numbers.push(numeric_constants);
          numbers.push(numbers.reduce((a, b) => a + b, 0)); // the sum of all numeric constants
          const dependences_wo_constants = ExcelUtils.all_cell_dependencies(
            formula,
            formulaY + 1,
            formulaX + 1,
            false
          );
          dependence_count.push(dependences_wo_constants.length);
          const r1c1 = ExcelUtils.formulaToR1C1(
            formula,
            formulaY + 1,
            formulaX + 1
          );
          const preface =
            ExcelUtils.column_index_to_name(formulaY + 1) +
            (formulaX + 1) +
            ':';
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
        const totalNumericDiff = Math.abs(numbers[0] - numbers[1]);

        // Omit fixes that are too small (too few cells).
        const fixRange = Colorize.expand(
          initial_adjusted_fixes[ind][1][0],
          initial_adjusted_fixes[ind][1][1]
        ).concat(
          Colorize.expand(
            initial_adjusted_fixes[ind][2][0],
            initial_adjusted_fixes[ind][2][1]
          )
        );
        if (fixRange.length < Colorize.minFixSize) {
          console.warn(
            'Omitted ' + JSON.stringify(print_formulas) + '(too small)'
          );
          continue;
        }

        // Entropy cutoff.
        const leftFixSize = Colorize.expand(
          initial_adjusted_fixes[ind][1][0],
          initial_adjusted_fixes[ind][1][1]
        ).length;
        const rightFixSize = Colorize.expand(
          initial_adjusted_fixes[ind][2][0],
          initial_adjusted_fixes[ind][2][1]
        ).length;
        const totalSize = leftFixSize + rightFixSize;
        const fixEntropy = -(
          (leftFixSize / totalSize) * Math.log2(leftFixSize / totalSize) +
          (rightFixSize / totalSize) * Math.log2(rightFixSize / totalSize)
        );
          // console.warn('fix entropy = ' + fixEntropy);
        if (fixEntropy > Colorize.maxEntropy) {
          console.warn(
            'Omitted ' + JSON.stringify(print_formulas) + '(too high entropy)'
          );
          continue;
        }

        // Binning.
        let bin = [];

        // Check for "fat" fixes (that result in more than a single row or single column).
        // Check if all in the same row.
        let sameRow = false;
        let sameColumn = false;
        {
          const fixColumn = initial_adjusted_fixes[ind][1][0][0];
          if (
            initial_adjusted_fixes[ind][1][1][0] === fixColumn &&
            initial_adjusted_fixes[ind][2][0][0] === fixColumn &&
            initial_adjusted_fixes[ind][2][1][0] === fixColumn
          ) {
            sameColumn = true;
          }
          const fixRow = initial_adjusted_fixes[ind][1][0][1];
          if (
            initial_adjusted_fixes[ind][1][1][1] === fixRow &&
            initial_adjusted_fixes[ind][2][0][1] === fixRow &&
            initial_adjusted_fixes[ind][2][1][1] === fixRow
          ) {
            sameRow = true;
          }
          if (!sameColumn && !sameRow) {
            bin.push(Colorize.BinCategories.FatFix);
          }
        }

        // Check for recurrent formulas. NOTE: not sure if this is working currently
        for (let i = 0; i < dependence_vectors.length; i++) {
          // If there are at least two dependencies and one of them is -1 in the column (row),
          // it is a recurrence (the common recurrence relation of starting at a value and
          // referencing, say, =B10+1).
          if (dependence_vectors[i].length > 0) {
            if (
              direction === 'vertical' &&
              dependence_vectors[i][0][0] === 0 &&
              dependence_vectors[i][0][1] === -1
            ) {
              bin.push(Colorize.BinCategories.RecurrentFormula);
              break;
            }
            if (
              direction === 'horizontal' &&
              dependence_vectors[i][0][0] === -1 &&
              dependence_vectors[i][0][1] === 0
            ) {
              bin.push(Colorize.BinCategories.RecurrentFormula);
              break;
            }
          }
        }
        // Different number of referents (dependencies).
        if (dependence_count[0] !== dependence_count[1]) {
          bin.push(Colorize.BinCategories.DifferentReferentCount);
        }
        // Different number of constants.
        if (all_numbers[0].length !== all_numbers[1].length) {
          if (Math.abs(all_numbers[0].length - all_numbers[1].length) === 1) {
            bin.push(Colorize.BinCategories.OneExtraConstant);
          } else {
            bin.push(Colorize.BinCategories.NumberOfConstantsMismatch);
          }
        }
        // Both constants.
        if (all_numbers[0].length > 0 && all_numbers[1].length > 0) {
          // Both have numbers.
          if (dependence_count[0] + dependence_count[1] === 0) {
            // Both have no dependents.
            bin.push(Colorize.BinCategories.BothConstants);
          } else {
            if (dependence_count[0] * dependence_count[1] === 0) {
              // One is a constant.
              bin.push(Colorize.BinCategories.OneIsAllConstants);
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
          if (
            JSON.stringify(dependence_vectors[0].sort()) !==
            JSON.stringify(dependence_vectors[1].sort())
          ) {
            bin.push(Colorize.BinCategories.R1C1Mismatch);
          }
        }
        // Different number of absolute ($, a.k.a. "anchor") references.
        if (absolute_refs[0] !== absolute_refs[1]) {
          bin.push(Colorize.BinCategories.AbsoluteRefMismatch);
        }
        // Dependencies that are neither vertical or horizontal (likely errors if an absolute-ref-mismatch).
        for (let i = 0; i < dependence_vectors.length; i++) {
          if (dependence_vectors[i].length > 0) {
            if (
              dependence_vectors[i][0][0] * dependence_vectors[i][0][1] !==
              0
            ) {
              bin.push(Colorize.BinCategories.OffAxisReference);
              break;
            }
          }
        }
        if (bin === []) {
          bin.push(Colorize.BinCategories.Unclassified);
        }
	  // In case there's more than one classification, prune some by priority (best explanation).
	  if (bin.includes(Colorize.BinCategories.OneIsAllConstants)) {
	      bin = [Colorize.BinCategories.OneIsAllConstants];
	  }
        // IMPORTANT:
        // Exclude reported bugs subject to certain conditions.
        if (
          bin.length > Colorize.maxCategories || // Too many categories
          (bin.indexOf(Colorize.BinCategories.FatFix) !== -1 &&
            Colorize.suppressFatFix) ||
          (bin.indexOf(Colorize.BinCategories.DifferentReferentCount) !== -1 &&
            Colorize.suppressDifferentReferentCount) ||
          (bin.indexOf(Colorize.BinCategories.RecurrentFormula) !== -1 &&
            Colorize.suppressRecurrentFormula) ||
          (bin.indexOf(Colorize.BinCategories.OneExtraConstant) !== -1 &&
            Colorize.suppressOneExtraConstant) ||
          (bin.indexOf(Colorize.BinCategories.NumberOfConstantsMismatch) !=
            -1 &&
            Colorize.suppressNumberOfConstantsMismatch) ||
          (bin.indexOf(Colorize.BinCategories.BothConstants) !== -1 &&
            Colorize.suppressBothConstants) ||
          (bin.indexOf(Colorize.BinCategories.OneIsAllConstants) !== -1 &&
            Colorize.suppressOneIsAllConstants) ||
          (bin.indexOf(Colorize.BinCategories.R1C1Mismatch) !== -1 &&
            Colorize.suppressR1C1Mismatch) ||
          (bin.indexOf(Colorize.BinCategories.AbsoluteRefMismatch) !== -1 &&
            Colorize.suppressAbsoluteRefMismatch) ||
          (bin.indexOf(Colorize.BinCategories.OffAxisReference) !== -1 &&
            Colorize.suppressOffAxisReference)
        ) {
          console.warn(
            'Omitted ' +
              JSON.stringify(print_formulas) +
              '(' +
              JSON.stringify(bin) +
              ')'
          );
          continue;
        } else {
          console.warn(
            'NOT omitted ' +
              JSON.stringify(print_formulas) +
              '(' +
              JSON.stringify(bin) +
              ')'
          );
        }
        final_adjusted_fixes.push(initial_adjusted_fixes[ind]);

        example_fixes_r1c1.push({
          bin: bin,
          direction: direction,
          numbers: numbers,
          numeric_difference: totalNumericDiff,
          magnitude_numeric_difference:
            totalNumericDiff === 0 ? 0 : Math.log10(totalNumericDiff),
          formulas: print_formulas,
          r1c1formulas: r1c1_print_formulas,
          dependence_vectors: dependence_vectors,
        });
        // example_fixes_r1c1.push([direction, formulas]);
      }

      let elapsed = myTimer.elapsedTime();
      if (Colorize.noElapsedTime) {
        elapsed = 0; // Dummy value, used for regression testing.
      }
      // Compute number of cells containing formulas.
      const numFormulaCells = sheet.formulas.flat().filter(x => x.length > 0)
        .length;

      // Count the number of non-empty cells.
      const numValueCells = sheet.values.flat().filter(x => x.length > 0)
        .length;

      // Compute total number of cells in the sheet (rows * columns).
      const columns = sheet.values[0].length;
      const rows = sheet.values.length;
      const totalCells = rows * columns;

      const out = {
        anomalousnessThreshold: Colorize.reportingThreshold,
        formattingDiscount: Colorize.formattingDiscount,
        // 'proposedFixes': final_adjusted_fixes,
        exampleFixes: example_fixes_r1c1,
        //		'exampleFixesR1C1' : example_fixes_r1c1,
        anomalousRanges: final_adjusted_fixes.length,
        weightedAnomalousRanges: 0, // actually calculated below.
        anomalousCells: 0, // actually calculated below.
        elapsedTimeSeconds: elapsed / 1e6,
        columns: columns,
        rows: rows,
        totalCells: totalCells,
        numFormulaCells: numFormulaCells,
        numValueCells: numValueCells,
      };

      // Compute precision and recall of proposed fixes, if we have annotated ground truth.
      const workbookBasename = path.basename(inp['workbookName']);
      // Build list of bugs.
      let foundBugs: any = final_adjusted_fixes.map(x => {
        if (x[0] >= Colorize.reportingThreshold / 100) {
          return Colorize.expand(x[1][0], x[1][1]).concat(
            Colorize.expand(x[2][0], x[2][1])
          );
        } else {
          return [];
        }
      });
      const foundBugsArray: any = Array.from(
        new Set(foundBugs.flat(1).map(JSON.stringify))
      );
      foundBugs = foundBugsArray.map(JSON.parse);
      out['anomalousCells'] = foundBugs.length;
      const weightedAnomalousRanges = final_adjusted_fixes
        .map(x => x[0])
        .reduce((x, y) => x + y, 0);
      out['weightedAnomalousRanges'] = weightedAnomalousRanges;

      /* disabled for now.

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
			sheetTruePositiveSet.add(workbookBasename + ':' + sheet.sheetName);
		    }
		    if (falsePositives.length) {
			sheetFalsePositives += 1;
			sheetFalsePositiveSet.add(workbookBasename + ':' + sheet.sheetName);
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
            */
      out['proposedFixes'] = final_adjusted_fixes;
      output.worksheets[sheet.sheetName] = out;
    }
    //	outputs.push(output);
    return output; // , scores, sheetTruePositiveSet, sheetTruePositives, sheetFalsePositiveSet, sheetFalsePositives };
  }

  // Convert a rectangle into a list of indices.
  public static expand(
    first: Colorize.excelintVector,
    second: Colorize.excelintVector
  ): Array<Colorize.excelintVector> {
    const [fcol, frow] = first;
    const [scol, srow] = second;
    const expanded: Array<Colorize.excelintVector> = [];
    for (let i = fcol; i <= scol; i++) {
      for (let j = frow; j <= srow; j++) {
        expanded.push([i, j, 0]);
      }
    }
    return expanded;
  }

  // Generate dependence vectors and their hash for all formulas.
  public static process_formulas(
    formulas: Array<Array<string>>,
    origin_col: number,
    origin_row: number
  ): Array<[Colorize.excelintVector, string]> {
    const base_vector = JSON.stringify(ExcelUtils.baseVector());
    const reducer = (
      acc: [number, number, number],
      curr: [number, number, number]
    ): [number, number, number] => [
      acc[0] + curr[0],
      acc[1] + curr[1],
      acc[2] + curr[2],
    ];
    const output: Array<[[number, number, number], string]> = [];

    // Compute the vectors for all of the formulas.
    for (let i = 0; i < formulas.length; i++) {
      const row = formulas[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        // If it's a formula, process it.
        if (cell.length > 0) {
          // FIXME MAYBE  && (row[j][0] === '=')) {
          const vec_array = ExcelUtils.all_dependencies(
            i,
            j,
            origin_row + i,
            origin_col + j,
            formulas
          );
          const adjustedX = j + origin_col + 1;
          const adjustedY = i + origin_row + 1;
          if (vec_array.length === 0) {
            if (cell[0] === '=') {
              // It's a formula but it has no dependencies (i.e., it just has constants). Use a distinguished value.
              output.push([
                [adjustedX, adjustedY, 0],
                Colorize.distinguishedZeroHash,
              ]);
            }
          } else {
            const vec = vec_array.reduce(reducer);
            if (JSON.stringify(vec) === base_vector) {
              // No dependencies! Use a distinguished value.
              // FIXME RESTORE THIS output.push([[adjustedX, adjustedY, 0], Colorize.distinguishedZeroHash]);
              output.push([[adjustedX, adjustedY, 0], vec.toString()]);
            } else {
              const hash = this.hash_vector(vec);
              const str = hash.toString();
              //			    console.log('hash for ' + adjustedX + ', ' + adjustedY + ' = ' + str);
              output.push([[adjustedX, adjustedY, 0], str]);
            }
          }
        }
      }
    }
    return output;
  }

  // Returns all referenced data so it can be colored later.
  public static color_all_data(refs: {
    [dep: string]: Array<Colorize.excelintVector>;
  }): Array<[Colorize.excelintVector, string]> {
    //	let t = new Timer('color_all_data');
    const referenced_data = [];
    for (const refvec of Object.keys(refs)) {
      const rv = refvec.split(',');
      const row = Number(rv[0]);
      const col = Number(rv[1]);
      referenced_data.push([[row, col, 0], Colorize.distinguishedZeroHash]); // See comment at top of function declaration.
    }
    //	t.split('processed all data');
    return referenced_data;
  }

  // Take all values and return an array of each row and column.
  // Note that for now, the last value of each tuple is set to 1.
  public static process_values(
    values: Array<Array<string>>,
    formulas: Array<Array<string>>,
    origin_col: number,
    origin_row: number
  ): Array<[Colorize.excelintVector, string]> {
    const value_array = [];
    //	let t = new Timer('process_values');
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        // If the value is not from a formula, include it.
        if (cell.length > 0 && formulas[i][j][0] !== '=') {
          const cellAsNumber = Number(cell).toString();
          if (cellAsNumber === cell) {
            // It's a number. Add it.
            const adjustedX = j + origin_col + 1;
            const adjustedY = i + origin_row + 1;
            value_array.push([
              [adjustedX, adjustedY, 1],
              Colorize.distinguishedZeroHash,
            ]); // See comment at top of function declaration.
          }
        }
      }
    }
    //	t.split('processed all values');
    return value_array;
  }

  // Take in a list of [[row, col], color] pairs and group them,
  // sorting them (e.g., by columns).
  private static identify_ranges(
    list: Array<[Colorize.excelintVector, string]>,
    sortfn?: (
      n1: Colorize.excelintVector,
      n2: Colorize.excelintVector
    ) => number
  ): {[val: string]: Array<Colorize.excelintVector>} {
    // Separate into groups based on their string value.
    const groups = {};
    for (const r of list) {
      groups[r[1]] = groups[r[1]] || [];
      groups[r[1]].push(r[0]);
    }
    // Now sort them all.
    for (const k of Object.keys(groups)) {
      groups[k].sort(sortfn);
    }
    return groups;
  }

  // Group all ranges by their value.
  private static group_ranges(
    groups: {[val: string]: Array<Colorize.excelintVector>},
    columnFirst: boolean
  ): {
    [val: string]: Array<[Colorize.excelintVector, Colorize.excelintVector]>;
  } {
    const output = {};
    let index0 = 0; // column
    let index1 = 1; // row
    if (!columnFirst) {
      index0 = 1; // row
      index1 = 0; // column
    }
    for (const k of Object.keys(groups)) {
      output[k] = [];
      let prev = groups[k].shift();
      let last = prev;
      for (const v of groups[k]) {
        // Check if in the same column, adjacent row (if columnFirst; otherwise, vice versa).
        if (v[index0] === last[index0] && v[index1] === last[index1] + 1) {
          last = v;
        } else {
          output[k].push([prev, last]);
          prev = v;
          last = v;
        }
      }
      output[k].push([prev, last]);
    }
    return output;
  }

  public static identify_groups(
    theList: Array<[Colorize.excelintVector, string]>
  ): {
    [val: string]: Array<[Colorize.excelintVector, Colorize.excelintVector]>;
  } {
    const columnsort = (
      a: Colorize.excelintVector,
      b: Colorize.excelintVector
    ) => {
      if (a[0] === b[0]) {
        return a[1] - b[1];
      } else {
        return a[0] - b[0];
      }
    };
    const id = this.identify_ranges(theList, columnsort);
    const gr = this.group_ranges(id, true); // column-first
    // Now try to merge stuff with the same hash.
    const newGr1 = JSONclone.clone(gr);
    const mg = this.merge_groups(newGr1);
    return mg;
  }

  public static processed_to_matrix(
    cols: number,
    rows: number,
    origin_col: number,
    origin_row: number,
    processed: Array<[Colorize.excelintVector, string]>
  ): Array<Array<number>> {
    // Invert the hash table.
    // First, initialize a zero-filled matrix.
    const matrix = new Array(cols);
    for (let i = 0; i < cols; i++) {
      matrix[i] = new Array(rows).fill(0);
    }
    // Now iterate through the processed formulas and update the matrix.
    for (const item of processed) {
      const [[col, row, isConstant], val] = item;
      // Yes, I know this is confusing. Will fix later.
      //	    console.log('C) cols = ' + rows + ', rows = ' + cols + '; row = ' + row + ', col = ' + col);
      const adjustedX = row - origin_row - 1;
      const adjustedY = col - origin_col - 1;
      let value = Number(Colorize.distinguishedZeroHash);
      if (isConstant === 1) {
        // That means it was a constant.
        // Set to a fixed value (as above).
      } else {
        value = Number(val);
      }
      matrix[adjustedX][adjustedY] = value;
    }
    return matrix;
  }

  public static stencilize(matrix: Array<Array<number>>): Array<Array<number>> {
    const stencil = Stencil.stencil_computation(
      matrix,
      (x: number, y: number) => {
        return x * y;
      },
      1
    );
    return stencil;
  }

  public static compute_stencil_probabilities(
    cols: number,
    rows: number,
    stencil: Array<Array<number>>
  ): Array<Array<number>> {
    //        console.log('compute_stencil_probabilities: stencil = ' + JSON.stringify(stencil));
    const probs = new Array(cols);
    for (let i = 0; i < cols; i++) {
      probs[i] = new Array(rows).fill(0);
    }
    // Generate the counts.
    let totalNonzeroes = 0;
    const counts = {};
    for (let i = 0; i < cols; i++) {
      for (let j = 0; j < rows; j++) {
        counts[stencil[i][j]] = counts[stencil[i][j]] + 1 || 1;
        if (stencil[i][j] !== 0) {
          totalNonzeroes += 1;
        }
      }
    }

    // Now iterate over the counts to compute probabilities.
    for (let i = 0; i < cols; i++) {
      for (let j = 0; j < rows; j++) {
        probs[i][j] = counts[stencil[i][j]] / totalNonzeroes;
      }
    }

    //	    console.log('probs = ' + JSON.stringify(probs));

    let totalEntropy = 0;
    let total = 0;
    for (let i = 0; i < cols; i++) {
      for (let j = 0; j < rows; j++) {
        if (stencil[i][j] > 0) {
          total += counts[stencil[i][j]];
        }
      }
    }

    for (let i = 0; i < cols; i++) {
      for (let j = 0; j < rows; j++) {
        if (counts[stencil[i][j]] > 0) {
          totalEntropy += this.entropy(counts[stencil[i][j]] / total);
        }
      }
    }

    const normalizedEntropy = totalEntropy / Math.log2(totalNonzeroes);

    // Now discount the probabilities by weighing them by the normalized total entropy.
    if (false) {
      for (let i = 0; i < cols; i++) {
        for (let j = 0; j < rows; j++) {
          probs[i][j] *= normalizedEntropy;
          //			totalEntropy += this.entropy(probs[i][j]);
        }
      }
    }
    return probs;
  }

  public static generate_suspicious_cells(
    cols: number,
    rows: number,
    origin_col: number,
    origin_row: number,
    matrix: Array<Array<number>>,
    probs: Array<Array<number>>,
    threshold = 0.01
  ): Array<Colorize.excelintVector> {
    const cells = [];
    let sumValues = 0;
    let countValues = 0;
    for (let i = 0; i < cols; i++) {
      for (let j = 0; j < rows; j++) {
        const adjustedX = j + origin_col + 1;
        const adjustedY = i + origin_row + 1;
        //		    console.log('examining ' + i + ' ' + j + ' = ' + matrix[i][j] + ' (' + adjustedX + ', ' + adjustedY + ')');
        if (probs[i][j] > 0) {
          sumValues += matrix[i][j];
          countValues += 1;
          if (probs[i][j] <= threshold) {
            // console.log('found one at ' + i + ' ' + j + ' = [' + matrix[i][j] + '] (' + adjustedX + ', ' + adjustedY + '): p = ' + probs[i][j]);
            if (matrix[i][j] !== 0) {
              // console.log('PUSHED!');
              // Never push an empty cell.
              cells.push([adjustedX, adjustedY, probs[i][j]]);
            }
          }
        }
      }
    }
    const avgValues = sumValues / countValues;
    cells.sort((a, b) => {
      return Math.abs(b[2] - avgValues) - Math.abs(a[2] - avgValues);
    });
    //        console.log('cells = ' + JSON.stringify(cells));
    return cells;
  }

  public static process_suspicious(
    usedRangeAddress: string,
    formulas: Array<Array<string>>,
    values: Array<Array<string>>
  ): [any, any, any, any] {
    if (false) {
      console.log('process_suspicious:');
      console.log(JSON.stringify(usedRangeAddress));
      console.log(JSON.stringify(formulas));
      console.log(JSON.stringify(values));
    }

    const t = new Timer('process_suspicious');

    const [sheetName, startCell] = ExcelUtils.extract_sheet_cell(
      usedRangeAddress
    );
    const origin = ExcelUtils.cell_dependency(startCell, 0, 0);

    let processed_formulas = [];
    // Filter out non-empty items from whole matrix.
    const totalFormulas = (formulas as any).flat().filter(Boolean).length;

    if (totalFormulas > this.formulasThreshold) {
      console.warn('Too many formulas to perform formula analysis.');
    } else {
      //	    t.split('about to process formulas');
      processed_formulas = Colorize.process_formulas(
        formulas,
        origin[0] - 1,
        origin[1] - 1
      );
      //	    t.split('processed formulas');
    }
    const useTimeouts = false;

    let referenced_data = [];
    let data_values = [];
    const cols = values.length;
    const rows = values[0].length;

    // Filter out non-empty items from whole matrix.
    const totalValues = (values as any).flat().filter(Boolean).length;
    if (totalValues > this.valuesThreshold) {
      console.warn('Too many values to perform reference analysis.');
    } else {
      // Compute references (to color referenced data).
      const refs = ExcelUtils.generate_all_references(
        formulas,
        origin[0] - 1,
        origin[1] - 1
      );
      //	    t.split('generated all references');

      referenced_data = Colorize.color_all_data(refs);
      // console.log('referenced_data = ' + JSON.stringify(referenced_data));
      data_values = Colorize.process_values(
        values,
        formulas,
        origin[0] - 1,
        origin[1] - 1
      );

      // t.split('processed data');
    }

    const grouped_data = Colorize.identify_groups(referenced_data);

    //	t.split('identified groups');

    const grouped_formulas = Colorize.identify_groups(processed_formulas);
    //	t.split('grouped formulas');

    // Identify suspicious cells.
    let suspicious_cells = [];

    if (values.length < 10000) {
      // Disabled for now. FIXME
      //            suspicious_cells = Colorize.find_suspicious_cells(cols, rows, origin, formulas, processed_formulas, data_values, 1 - Colorize.getReportingThreshold() / 100); // Must be more rare than this fraction.
      suspicious_cells = Colorize.find_suspicious_cells(
        cols,
        rows,
        origin,
        formulas,
        processed_formulas,
        data_values,
        1
      ); // Must be more rare than this fraction.
    }

    const proposed_fixes = Colorize.generate_proposed_fixes(grouped_formulas);

    if (false) {
      console.log('results:');
      console.log(JSON.stringify(suspicious_cells));
      console.log(JSON.stringify(grouped_formulas));
      console.log(JSON.stringify(grouped_data));
      console.log(JSON.stringify(proposed_fixes));
    }

    return [suspicious_cells, grouped_formulas, grouped_data, proposed_fixes];
  }

  // Shannon entropy.
  public static entropy(p: number): number {
    return -p * Math.log2(p);
  }

  // Take two counts and compute the normalized entropy difference that would result if these were 'merged'.
  public static entropydiff(oldcount1, oldcount2) {
    const total = oldcount1 + oldcount2;
    const prevEntropy =
      this.entropy(oldcount1 / total) + this.entropy(oldcount2 / total);
    const normalizedEntropy = prevEntropy / Math.log2(total);
    return -normalizedEntropy;
  }

  // Compute the normalized distance from merging two ranges.
  public static fix_metric(
    target_norm: number,
    target: [Colorize.excelintVector, Colorize.excelintVector],
    merge_with_norm: number,
    merge_with: [Colorize.excelintVector, Colorize.excelintVector]
  ): number {
    //	console.log('fix_metric: ' + target_norm + ', ' + JSON.stringify(target) + ', ' + merge_with_norm + ', ' + JSON.stringify(merge_with));
    const [t1, t2] = target;
    const [m1, m2] = merge_with;
    const n_target = RectangleUtils.area([
      [t1[0], t1[1], 0],
      [t2[0], t2[1], 0],
    ]);
    const n_merge_with = RectangleUtils.area([
      [m1[0], m1[1], 0],
      [m2[0], m2[1], 0],
    ]);
    const n_min = Math.min(n_target, n_merge_with);
    const n_max = Math.max(n_target, n_merge_with);
    const norm_min = Math.min(merge_with_norm, target_norm);
    const norm_max = Math.max(merge_with_norm, target_norm);
    let fix_distance = Math.abs(norm_max - norm_min) / this.Multiplier;
    // Ensure that the minimum fix is at least one (we need this if we don't use the L1 norm).
    if (fix_distance < 1.0) {
      fix_distance = 1.0;
    }
    const entropy_drop = this.entropydiff(n_min, n_max); // negative
    let ranking = (1.0 + entropy_drop) / (fix_distance * n_min); // ENTROPY WEIGHTED BY FIX DISTANCE
    ranking = -ranking; // negating to sort in reverse order.
    return ranking;
  }

  // Iterate through the size of proposed fixes.
  public static count_proposed_fixes(
    fixes: Array<
      [
        number,
        [Colorize.excelintVector, Colorize.excelintVector],
        [Colorize.excelintVector, Colorize.excelintVector]
      ]
    >
  ): number {
    let count = 0;
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const [f11, f12] = fixes[k][1];
      const [f21, f22] = fixes[k][2];
      count += RectangleUtils.diagonal([
        [f11[0], f11[1], 0],
        [f12[0], f12[1], 0],
      ]);
      count += RectangleUtils.diagonal([
        [f21[0], f21[1], 0],
        [f22[0], f22[1], 0],
      ]);
    }
    return count;
  }

  // Try to merge fixes into larger groups.
  public static fix_proposed_fixes(
    fixes: Array<
      [
        number,
        [Colorize.excelintVector, Colorize.excelintVector],
        [Colorize.excelintVector, Colorize.excelintVector]
      ]
    >
  ): Array<
    [
      number,
      [Colorize.excelintVector, Colorize.excelintVector],
      [Colorize.excelintVector, Colorize.excelintVector]
    ]
  > {
    // example: [[-0.8729568798082977,[[4,23],[13,23]],[[3,23,0],[3,23,0]]],[-0.6890824929174288,[[4,6],[7,6]],[[3,6,0],[3,6,0]]],[-0.5943609377704335,[[4,10],[6,10]],[[3,10,0],[3,10,0]]],[-0.42061983571430495,[[3,27],[4,27]],[[5,27,0],[5,27,0]]],[-0.42061983571430495,[[4,14],[5,14]],[[3,14,0],[3,14,0]]],[-0.42061983571430495,[[6,27],[7,27]],[[5,27,0],[5,27,0]]]]
    const count = 0;
    // Search for fixes where the same coordinate pair appears in the front and in the back.
    const front = {};
    const back = {};
    // Build up the front and back dictionaries.
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      // Sort the fixes so the smaller array (further up and
      // to the left) always comes first.
      if (fixes[k][1] > fixes[k][2]) {
        const tmp = fixes[k][1];
        fixes[k][1] = fixes[k][2];
        fixes[k][2] = tmp;
      }
      // Now add them.
      front[JSON.stringify(fixes[k][1])] = fixes[k];
      back[JSON.stringify(fixes[k][2])] = fixes[k];
    }
    // Now iterate through one, looking for hits on the other.
    const new_fixes = [];
    const merged = {};
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const original_score = fixes[k][0];
      if (-original_score < Colorize.reportingThreshold / 100) {
        continue;
      }
      const this_front_str = JSON.stringify(fixes[k][1]);
      const this_back_str = JSON.stringify(fixes[k][2]);
      if (!(this_front_str in back) && !(this_back_str in front)) {
        // No match. Just merge them.
        new_fixes.push(fixes[k]);
      } else {
        if (!merged[this_front_str] && this_front_str in back) {
          // FIXME: does this score make sense? Verify mathematically.
          const newscore =
            -original_score * JSON.parse(back[this_front_str][0]);
          const new_fix = [newscore, fixes[k][1], back[this_front_str][1]];
          new_fixes.push(new_fix);
          merged[this_front_str] = true;
          // FIXME? testing below. The idea is to not keep merging things (for now).
          merged[this_back_str] = true;
          continue;
        }
        if (!merged[this_back_str] && this_back_str in front) {
          // this_back_str in front
          //			console.log('**** (2) merging ' + this_back_str + ' with ' + JSON.stringify(front[this_back_str]));
          // FIXME. This calculation may not make sense.
          const newscore =
            -original_score * JSON.parse(front[this_back_str][0]);
          //			console.log('pushing ' + JSON.stringify(fixes[k][1]) + ' with ' + JSON.stringify(front[this_back_str][1]));
          const new_fix = [newscore, fixes[k][1], front[this_back_str][2]];
          //			console.log('pushing ' + JSON.stringify(new_fix));
          new_fixes.push(new_fix);
          merged[this_back_str] = true;
          // FIXME? testing below.
          merged[this_front_str] = true;
        }
      }
    }
    return new_fixes;
  }

  public static generate_proposed_fixes(groups: {
    [val: string]: Array<[Colorize.excelintVector, Colorize.excelintVector]>;
  }): Array<
    [
      number,
      [Colorize.excelintVector, Colorize.excelintVector],
      [Colorize.excelintVector, Colorize.excelintVector]
    ]
  > {
    //	let t = new Timer('generate_proposed_fixes');
    //	t.split('about to find.');
    const proposed_fixes_new = find_all_proposed_fixes(groups);
    //	t.split('sorting fixes.');
    proposed_fixes_new.sort((a, b) => {
      return a[0] - b[0];
    });
    //	t.split('done.');
    //	console.log(JSON.stringify(proposed_fixes_new));
    return proposed_fixes_new;
  }

  public static merge_groups(groups: {
    [val: string]: Array<[Colorize.excelintVector, Colorize.excelintVector]>;
  }): {
    [val: string]: Array<[Colorize.excelintVector, Colorize.excelintVector]>;
  } {
    for (const k of Object.keys(groups)) {
      const g = groups[k].slice();
      groups[k] = this.merge_individual_groups(g);
    }
    return groups;
  }

  public static merge_individual_groups(
    group: Array<[Colorize.excelintVector, Colorize.excelintVector]>
  ): Array<[Colorize.excelintVector, Colorize.excelintVector]> {
    const t = new Timer('merge_individual_groups');
    let numIterations = 0;
    group = group.sort();
    while (true) {
      let merged_one = false;
      const deleted_rectangles = {};
      const updated_rectangles = [];
      const working_group = group.slice();
      while (working_group.length > 0) {
        const head = working_group.shift();
        for (let i = 0; i < working_group.length; i++) {
          if (RectangleUtils.is_mergeable(head, working_group[i])) {
            const head_str = JSON.stringify(head);
            const working_group_i_str = JSON.stringify(working_group[i]);
            // NB: 12/7/19 New check below, used to be unconditional.
            if (
              !(head_str in deleted_rectangles) &&
              !(working_group_i_str in deleted_rectangles)
            ) {
              updated_rectangles.push(
                RectangleUtils.bounding_box(head, working_group[i])
              );
              deleted_rectangles[head_str] = true;
              deleted_rectangles[working_group_i_str] = true;
              merged_one = true;
              break; // was disabled
            }
          }
        }
      }
      for (let i = 0; i < group.length; i++) {
        if (!(JSON.stringify(group[i]) in deleted_rectangles)) {
          updated_rectangles.push(group[i]);
        }
      }
      updated_rectangles.sort();
      if (!merged_one) {
        return updated_rectangles;
      }
      group = updated_rectangles.slice();
      numIterations++;
      if (numIterations > 2000) {
        // This is a hack to guarantee convergence.
        console.log('Too many iterations; abandoning this group.');
        t.split('done, ' + numIterations + ' iterations.');
        return [
          [
            [-1, -1, 0],
            [-1, -1, 0],
          ],
        ];
      }
    }
  }

  public static hash_vector(vec: Array<number>): number {
    const useL1norm = true; // false;
    if (useL1norm) {
      const baseX = 0; // 7;
      const baseY = 0; // 3;
      const v0 = Math.abs(vec[0] - baseX);
      //	v0 = v0 * v0;
      const v1 = Math.abs(vec[1] - baseY);
      //	v1 = v1 * v1;
      const v2 = vec[2];
      return this.Multiplier * (v0 + v1 + v2);
    } else {
      const baseX = -7; // was 7
      const baseY = -3; // was 3
      let v0 = vec[0] - baseX;
      v0 = v0 * v0;
      let v1 = vec[1] - baseY;
      v1 = v1 * v1;
      const v2 = vec[2];
      return this.Multiplier * Math.sqrt(v0 + v1 + v2);
    }
    //	return this.Multiplier * (Math.sqrt(v0 + v1) + v2);
  }

  // Discount proposed fixes if they have different formats.
  public static adjust_proposed_fixes(
    fixes: any[],
    propertiesToGet: any[][],
    origin_col: number,
    origin_row: number
  ): any {
    const proposed_fixes = [];
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      // Format of proposed fixes =, e.g., [-3.016844756293869, [[5,7],[5,11]],[[6,7],[6,11]]]
      // entropy, and two ranges:
      //    upper-left corner of range (column, row), lower-right corner of range (column, row)

      const score = fixes[k][0];
      // Get rid of scores below 0.01.
      if (-score * 100 < 1) {
        // console.log('trimmed ' + (-score));
        continue;
      }
      // Sort the fixes.
      // This is a pain because if we don't pad appropriately, [1,9] is 'less than' [1,10]. (Seriously.)
      // So we make sure that numbers are always left padded with zeroes to make the number 10 digits long
      // (which is 1 more than Excel needs right now).
      const firstPadded = fixes[k][1].map(a => a.toString().padStart(10, '0'));
      const secondPadded = fixes[k][2].map(a => a.toString().padStart(10, '0'));

      const first = firstPadded < secondPadded ? fixes[k][1] : fixes[k][2];
      const second = firstPadded < secondPadded ? fixes[k][2] : fixes[k][1];

      const [[ax1, ay1], [ax2, ay2]] = first;
      const [[bx1, by1], [bx2, by2]] = second;

      const col0 = ax1 - origin_col - 1;
      const row0 = ay1 - origin_row - 1;
      const col1 = bx2 - origin_col - 1;
      const row1 = by2 - origin_row - 1;

      // Now check whether the formats are all the same or not.
      let sameFormats = true;
      const firstFormat = JSON.stringify(propertiesToGet[row0][col0]);
      for (let i = row0; i <= row1; i++) {
        for (let j = col0; j <= col1; j++) {
          const str = JSON.stringify(propertiesToGet[i][j]);
          if (str !== firstFormat) {
            sameFormats = false;
            break;
          }
        }
      }
      proposed_fixes.push([score, first, second, sameFormats]);
    }
    return proposed_fixes;
  }

  public static find_suspicious_cells(
    cols: number,
    rows: number,
    origin: [number, number, number],
    formulas: any[][],
    processed_formulas: any[],
    data_values: any,
    threshold: number
  ) {
    return []; // FIXME disabled for now
    let suspiciousCells: any[];
      {
      // data_values = data_values;
      const formula_matrix = Colorize.processed_to_matrix(
        cols,
        rows,
        origin[0] - 1,
        origin[1] - 1,
        //								processed_formulas);
        processed_formulas.concat(data_values)
      );
      //									processed_formulas);

      const stencil = Colorize.stencilize(formula_matrix);
     // console.log('after stencilize: stencil = ' + JSON.stringify(stencil));
      const probs = Colorize.compute_stencil_probabilities(cols, rows, stencil);
      // console.log('probs = ' + JSON.stringify(probs));

      const candidateSuspiciousCells = Colorize.generate_suspicious_cells(
        cols,
        rows,
        origin[0] - 1,
        origin[1] - 1,
        formula_matrix,
        probs,
        threshold
      );
      // Prune any cell that is in fact a formula.

      if (typeof formulas !== 'undefined') {
        let totalFormulaWeight = 0;
        let totalWeight = 0;
        suspiciousCells = candidateSuspiciousCells.filter(c => {
          const theFormula = formulas[c[1] - origin[1]][c[0] - origin[0]];
          if (theFormula.length < 1 || theFormula[0] !== '=') {
            totalWeight += c[2];
            return true;
          } else {
            // It's a formula: we will remove it, but also track how much it contributed to the probability distribution.
            totalFormulaWeight += c[2];
            totalWeight += c[2];
            return false;
          }
        });
        // console.log('total formula weight = ' + totalFormulaWeight);
        // console.log('total weight = ' + totalWeight);
        // Now we need to correct all the non-formulas to give them weight proportional to the case when the formulas are removed.
        const multiplier = totalFormulaWeight / totalWeight;
        console.log(
          'after processing 1, suspiciousCells = ' +
            JSON.stringify(suspiciousCells)
        );
        suspiciousCells = suspiciousCells.map(c => [
          c[0],
          c[1],
          c[2] * multiplier,
        ]);
        console.log(
          'after processing 2, suspiciousCells = ' +
            JSON.stringify(suspiciousCells)
        );
        suspiciousCells = suspiciousCells.filter(c => c[2] <= threshold);
        console.log(
          'after processing 3, suspiciousCells = ' +
            JSON.stringify(suspiciousCells)
        );
      } else {
        suspiciousCells = candidateSuspiciousCells;
      }
    }
    return suspiciousCells;
  }
}

export namespace Colorize {
  export type excelintVector = [number, number, number];

  export enum BinCategories {
    FatFix = 'Inconsistent multiple columns/rows', // fix is not a single column or single row
    RecurrentFormula = 'Formula(s) refer to each other', // formulas refer to each other
    OneExtraConstant = 'Formula(s) with an extra constant', // one has no constant and the other has one constant
    NumberOfConstantsMismatch = 'Formulas have different number of constants', // both have constants but not the same number of constants
    BothConstants = 'All constants, but different values', // both have only constants but differ in numeric value
    OneIsAllConstants = 'Mix of constants and formulas', // one is entirely constants and other is formula
    AbsoluteRefMismatch = 'Mix of absolute ($) and regular references', // relative vs. absolute mismatch
    OffAxisReference = 'References refer to different rows/columns', // references refer to different columns or rows
    R1C1Mismatch = 'Refers to different ranges', // different R1C1 representations
    DifferentReferentCount = 'Formula ranges are of different sizes', // ranges have different number of referents
    // Not yet implemented.
    RefersToEmptyCells = 'Formulas refer to empty cells',
    UsesDifferentOperations = 'Formulas use different functions', // e.g. SUM vs. AVERAGE
    // Fall-through category
    Unclassified = 'unclassified',
  }
}
