// Polyfill for flat (IE & Edge)
const flat = require("array.prototype.flat");
flat.shim();

import { ExcelUtils } from "./excelutils";
import { RectangleUtils } from "./rectangleutils";
import { Timer } from "./timer";
import { JSONclone } from "./jsonclone";
import { find_all_proposed_fixes } from "./groupme";
import { Stencil } from "./infogain";
import * as XLNT from "./ExceLintTypes";
import { Dict } from "./ExceLintTypes";
import { WorkbookOutput, WorksheetOutput } from "./exceljson";
import { Config } from "./config";
import { Classification } from "./classification";

export class Colorize {
  // Color-blind friendly color palette.
  public static palette = [
    "#ecaaae",
    "#74aff3",
    "#d8e9b2",
    "#deb1e0",
    "#9ec991",
    "#adbce9",
    "#e9c59a",
    "#71cdeb",
    "#bfbb8a",
    "#94d9df",
    "#91c7a8",
    "#b4efd3",
    "#80b6aa",
    "#9bd1c6",
  ]; // removed '#73dad1'

  // True iff this class been initialized.
  private static initialized = false;

  // The array of colors (used to hash into).
  private static color_list = [];

  // A hash string indicating no dependencies; in other words,
  // either a formula that makes no references (like `=RAND()`) or a data cell (like `1`)
  private static noDependenciesHash = "12345";

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

  // return true if this sheet is not the same as the other sheet
  public static isNotSameSheet(thisSheetName: string, otherSheetName: string): boolean {
    return thisSheetName !== "" && otherSheetName !== thisSheetName;
  }

  // returns true if this is an empty sheet
  public static isEmptySheet(sheet: any): boolean {
    return sheet.formulas.length === 0 && sheet.values.length === 0;
  }

  // Get rid of multiple exclamation points in the used range address,
  // as these interfere with later regexp parsing.
  public static normalizeAddress(addr: string): string {
    return addr.replace(/!(!+)/, "!");
  }

  // Filter fixes by entropy score threshold
  private static filterFixesByUserThreshold(
    fixes: XLNT.ProposedFix[],
    thresh: number
  ): XLNT.ProposedFix[] {
    const fixes2: XLNT.ProposedFix[] = [];
    for (let ind = 0; ind < fixes.length; ind++) {
      const pf = fixes[ind];
      let adjusted_score = -pf.score;
      if (adjusted_score * 100 >= thresh) {
        fixes2.push(new XLNT.ProposedFix(adjusted_score, pf.rect1, pf.rect2));
      }
    }
    return fixes2;
  }

  // Returns true if the "direction" of a fix is vertical
  private static fixIsVertical(fix: XLNT.ProposedFix): boolean {
    const rect1_ul_x = XLNT.upperleft(fix.rect1).x;
    const rect2_ul_x = XLNT.upperleft(fix.rect2).x;
    return rect1_ul_x === rect2_ul_x;
  }

  private static fixCellCount(fix: XLNT.ProposedFix): number {
    const fixRange = XLNT.expand(XLNT.upperleft(fix.rect1), XLNT.bottomright(fix.rect1)).concat(
      XLNT.expand(XLNT.upperleft(fix.rect2), XLNT.bottomright(fix.rect2))
    );
    return fixRange.length;
  }

  private static fixEntropy(fix: XLNT.ProposedFix): number {
    const leftFixSize = XLNT.expand(XLNT.upperleft(fix.rect1), XLNT.bottomright(fix.rect1)).length;
    const rightFixSize = XLNT.expand(XLNT.upperleft(fix.rect2), XLNT.bottomright(fix.rect2)).length;
    const totalSize = leftFixSize + rightFixSize;
    const fixEntropy = -(
      (leftFixSize / totalSize) * Math.log2(leftFixSize / totalSize) +
      (rightFixSize / totalSize) * Math.log2(rightFixSize / totalSize)
    );
    return fixEntropy;
  }

  // Performs an analysis on an entire workbook
  public static process_workbook(inp: WorkbookOutput, sheetName: string): XLNT.WorkbookAnalysis {
    const wba = new XLNT.WorkbookAnalysis();

    // look for the requested sheet
    for (let i = 0; i < inp.worksheets.length; i++) {
      const sheet = inp.worksheets[i];

      // skip sheets that don't match sheetName or are empty
      if (Colorize.isNotSameSheet(sheetName, sheet.sheetName) || Colorize.isEmptySheet(sheet)) {
        continue;
      }

      // get the used range
      const usedRangeAddress = Colorize.normalizeAddress(sheet.usedRangeAddress);

      // start timer
      const myTimer = new Timer("excelint");

      // Get anomalous cells and proposed fixes, among others.
      const a = Colorize.process_suspicious(usedRangeAddress, sheet.formulas, sheet.values);

      // Eliminate fixes below user threshold
      a.proposed_fixes = Colorize.filterFixesByUserThreshold(
        a.proposed_fixes,
        Config.reportingThreshold
      );

      // Remove fixes that require fixing both a formula AND formatting.
      // NB: origin_col and origin_row currently hard-coded at 0,0.
      Colorize.adjust_proposed_fixes(
        a.proposed_fixes,
        sheet.styles,
        0,
        0
      );

      // Process all the fixes, classifying and optionally pruning them.
      const final_adjusted_fixes: XLNT.ProposedFix[] = []; // We will eventually trim these.
      for (let ind = 0; ind < a.proposed_fixes.length; ind++) {
        // Get this fix
        const fix = a.proposed_fixes[ind];

        // Determine the direction of the range (vertical or horizontal) by looking at the axes.
        const is_vert: boolean = Colorize.fixIsVertical(fix);

        // Formula info for each rectangle
        const rect_info = fix.rectangles.map((rect) => new XLNT.RectInfo(rect, sheet));

        // Omit fixes that are too small (too few cells).
        if (Colorize.fixCellCount(fix) < Config.minFixSize) {
          const print_formulas = JSON.stringify(rect_info.map((fi) => fi.print_formula));
          console.warn("Omitted " + print_formulas + "(too small)");
          continue;
        }

        // Omit fixes with entropy change over threshold
        if (Colorize.fixEntropy(fix) > Config.maxEntropy) {
          const print_formulas = JSON.stringify(rect_info.map((fi) => fi.print_formula));
          console.warn("Omitted " + JSON.stringify(print_formulas) + "(too high entropy)");
          continue;
        }

        // Classify fixes & prune based on the best explanation
        const bin = Classification.pruneFixes(
          Classification.classifyFixes(fix, is_vert, rect_info)
        );

        // IMPORTANT:
        // Exclude reported bugs subject to certain conditions.
        if (Classification.omitFixes(bin, rect_info)) continue;

        // If we're still here, accept this fix
        final_adjusted_fixes.push(fix);

        // Package everything up with the fix
        fix.analysis = new XLNT.FixAnalysis(bin, rect_info, is_vert);
      }

      let elapsed = myTimer.elapsedTime();
      if (Config.noElapsedTime) {
        elapsed = 0; // Dummy value, used for regression testing.
      }

      // gather all statistics about the sheet
      wba.addSheet(new XLNT.WorksheetAnalysis(sheet, final_adjusted_fixes, a));
    }
    return wba;
  }

  // Generate dependence vectors and their hash for all formulas.
  public static process_formulas(
    formulas: XLNT.Spreadsheet,
    origin_col: number,
    origin_row: number
  ): [XLNT.ExceLintVector, XLNT.Fingerprint][] {
    const base_vector = ExcelUtils.baseVector();
    const output: Array<[XLNT.ExceLintVector, XLNT.Fingerprint]> = [];

    // Compute the vectors for all of the formulas.
    for (let i = 0; i < formulas.length; i++) {
      const row = formulas[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        // If it's a formula, process it.
        if (cell.length > 0) {
          // FIXME MAYBE  && (row[j][0] === '=')) {
          const vec_array: XLNT.ExceLintVector[] = ExcelUtils.all_dependencies(
            i,
            j,
            origin_row + i,
            origin_col + j,
            formulas
          );
          const adjustedX = j + origin_col + 1;
          const adjustedY = i + origin_row + 1;
          if (vec_array.length === 0) {
            if (cell[0] === "=") {
              // It's a formula but it has no dependencies (i.e., it just has constants). Use a distinguished value.
              output.push([
                new XLNT.ExceLintVector(adjustedX, adjustedY, 0),
                Colorize.noDependenciesHash,
              ]);
            }
          } else {
            const vec = vec_array.reduce(XLNT.ExceLintVector.VectorSum);
            if (vec.equals(base_vector)) {
              // No dependencies! Use a distinguished value.
              // Emery's FIXME: RESTORE THIS output.push([[adjustedX, adjustedY, 0], Colorize.distinguishedZeroHash]);
              // DAN TODO: I don't understand this case.
              output.push([
                new XLNT.ExceLintVector(adjustedX, adjustedY, 0),
                Colorize.noDependenciesHash,
              ]);
            } else {
              const hash = vec.hash();
              output.push([new XLNT.ExceLintVector(adjustedX, adjustedY, 0), hash.toString()]);
            }
          }
        }
      }
    }
    return output;
  }

  // Returns all referenced data so it can be colored later.
  public static color_all_data(refs: Dict<boolean>): [XLNT.ExceLintVector, XLNT.Fingerprint][] {
    const referenced_data: [XLNT.ExceLintVector, XLNT.Fingerprint][] = [];
    for (const refvec of Object.keys(refs)) {
      const rv = refvec.split(",");
      const row = Number(rv[0]);
      const col = Number(rv[1]);
      referenced_data.push([new XLNT.ExceLintVector(row, col, 0), Colorize.noDependenciesHash]); // See comment at top of function declaration.
    }
    return referenced_data;
  }

  // Take all values and return an array of each row and column.
  // Note that for now, the last value of each tuple is set to 1.
  public static process_values(
    values: XLNT.Spreadsheet,
    formulas: XLNT.Spreadsheet,
    origin_col: number,
    origin_row: number
  ): [XLNT.ExceLintVector, XLNT.Fingerprint][] {
    const value_array: [XLNT.ExceLintVector, XLNT.Fingerprint][] = [];
    //	let t = new Timer('process_values');
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        // If the value is not from a formula, include it.
        if (cell.length > 0 && formulas[i][j][0] !== "=") {
          const cellAsNumber = Number(cell).toString();
          if (cellAsNumber === cell) {
            // It's a number. Add it.
            const adjustedX = j + origin_col + 1;
            const adjustedY = i + origin_row + 1;
            // See comment at top of function declaration for DistinguishedZeroHash
            value_array.push([
              new XLNT.ExceLintVector(adjustedX, adjustedY, 1),
              Colorize.noDependenciesHash,
            ]);
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
    list: Array<[XLNT.ExceLintVector, string]>,
    sortfn: (n1: XLNT.ExceLintVector, n2: XLNT.ExceLintVector) => number
  ): Dict<XLNT.ExceLintVector[]> {
    // Separate into groups based on their XLNT.Fingerprint value.
    const groups = {};
    for (const r of list) {
      const [vec, fp] = r;
      groups[fp] = groups[fp] || []; // initialize array if necessary
      groups[fp].push(vec);
    }
    // Now sort them all.
    for (const k of Object.keys(groups)) {
      groups[k].sort(sortfn);
    }
    return groups;
  }

  // Collect all ranges of cells that share a XLNT.Fingerprint
  private static find_contiguous_regions(
    groups: Dict<XLNT.ExceLintVector[]>
  ): Dict<XLNT.Rectangle[]> {
    const output: Dict<XLNT.Rectangle[]> = {};

    for (const key of Object.keys(groups)) {
      // Here, we scan all of the vectors in this group, accumulating
      // all adjacent vectors by tracking the start and end. Whevener
      // we encounter a non-adjacent vector, push the region to the output
      // list and then start tracking a new region.
      output[key] = [];
      let start = groups[key].shift(); // remove the first vector from the list
      let end = start;
      for (const v of groups[key]) {
        // Check if v is in the same column as last, adjacent row
        if (v.x === end.x && v.y === end.y + 1) {
          end = v;
        } else {
          output[key].push([start, end]);
          start = v;
          end = v;
        }
      }
      output[key].push([start, end]);
    }
    return output;
  }

  public static identify_groups(theList: [XLNT.ExceLintVector, string][]): Dict<XLNT.Rectangle[]> {
    const id = Colorize.identify_ranges(theList, ExcelUtils.ColumnSort);
    const gr = Colorize.find_contiguous_regions(id);
    // Now try to merge stuff with the same hash.
    const newGr1 = JSONclone.clone(gr);
    const mg = Colorize.merge_groups(newGr1);
    return mg;
  }

  public static generate_suspicious_cells(
    cols: number,
    rows: number,
    origin_col: number,
    origin_row: number,
    matrix: Array<Array<number>>,
    probs: Array<Array<number>>,
    threshold = 0.01
  ): Array<XLNT.ExceLintVector> {
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
    formulas: XLNT.Spreadsheet,
    values: XLNT.Spreadsheet
  ): XLNT.Analysis {
    const t = new Timer("process_suspicious");

    const [sheetName, startCell] = ExcelUtils.extract_sheet_cell(usedRangeAddress);
    const origin = ExcelUtils.cell_dependency(startCell, 0, 0);

    let processed_formulas: [XLNT.ExceLintVector, string][] = [];
    // Filter out non-empty items from whole matrix.
    const totalFormulas = (formulas as any).flat().filter(Boolean).length;

    if (totalFormulas > Config.formulasThreshold) {
      console.warn("Too many formulas to perform formula analysis.");
    } else {
      processed_formulas = Colorize.process_formulas(formulas, origin.x - 1, origin.y - 1);
    }

    let referenced_data: [XLNT.ExceLintVector, XLNT.Fingerprint][] = [];
    let data_values: [XLNT.ExceLintVector, XLNT.Fingerprint][] = [];
    const cols = values.length;
    const rows = values[0].length;

    // Filter out non-empty items from whole matrix.
    const totalValues = (values as any).flat().filter(Boolean).length;
    if (totalValues > Config.valuesThreshold) {
      console.warn("Too many values to perform reference analysis.");
    } else {
      // Compute references (to color referenced data).
      const refs: Dict<boolean> = ExcelUtils.generate_all_references(
        formulas,
        origin.x - 1,
        origin.y - 1
      );

      referenced_data = Colorize.color_all_data(refs);
      data_values = Colorize.process_values(values, formulas, origin.x - 1, origin.y - 1);
    }

    // find regions for data
    const grouped_data = Colorize.identify_groups(referenced_data);

    // find regions for formulas
    const grouped_formulas = Colorize.identify_groups(processed_formulas);

    // Identify suspicious cells (disabled)
    let suspicious_cells: XLNT.ExceLintVector[] = [];

    // find proposed fixes
    const proposed_fixes = Colorize.generate_proposed_fixes(grouped_formulas);

    return new XLNT.Analysis(suspicious_cells, grouped_formulas, grouped_data, proposed_fixes);
  }

  // Shannon entropy.
  public static entropy(p: number): number {
    return -p * Math.log2(p);
  }

  // Take two counts and compute the normalized entropy difference that would result if these were 'merged'.
  public static entropydiff(oldcount1, oldcount2) {
    const total = oldcount1 + oldcount2;
    const prevEntropy = this.entropy(oldcount1 / total) + this.entropy(oldcount2 / total);
    const normalizedEntropy = prevEntropy / Math.log2(total);
    return -normalizedEntropy;
  }

  // Compute the normalized distance from merging two ranges.
  public static compute_fix_metric(
    target_norm: number,
    target: XLNT.Rectangle,
    merge_with_norm: number,
    merge_with: XLNT.Rectangle
  ): XLNT.Metric {
    //	console.log('fix_metric: ' + target_norm + ', ' + JSON.stringify(target) + ', ' + merge_with_norm + ', ' + JSON.stringify(merge_with));
    const [t1, t2] = target;
    const [m1, m2] = merge_with;
    const n_target = RectangleUtils.area([
      new XLNT.ExceLintVector(t1.x, t1.y, 0),
      new XLNT.ExceLintVector(t2.x, t2.y, 0),
    ]);
    const n_merge_with = RectangleUtils.area([
      new XLNT.ExceLintVector(m1.x, m1.y, 0),
      new XLNT.ExceLintVector(m2.x, m2.y, 0),
    ]);
    const n_min = Math.min(n_target, n_merge_with);
    const n_max = Math.max(n_target, n_merge_with);
    const norm_min = Math.min(merge_with_norm, target_norm);
    const norm_max = Math.max(merge_with_norm, target_norm);
    let fix_distance = Math.abs(norm_max - norm_min) / XLNT.ExceLintVector.Multiplier;

    // Ensure that the minimum fix is at least one (we need this if we don't use the L1 norm).
    if (fix_distance < 1.0) {
      fix_distance = 1.0;
    }
    const entropy_drop = this.entropydiff(n_min, n_max); // negative
    const ranking = (1.0 + entropy_drop) / (fix_distance * n_min); // ENTROPY WEIGHTED BY FIX DISTANCE
    return -ranking; // negating to sort in reverse order.
  }

  // Iterate through the size of proposed fixes.
  public static count_proposed_fixes(
    fixes: Array<
      [
        number,
        [XLNT.ExceLintVector, XLNT.ExceLintVector],
        [XLNT.ExceLintVector, XLNT.ExceLintVector]
      ]
    >
  ): number {
    let count = 0;
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const [f11, f12] = fixes[k][1];
      const [f21, f22] = fixes[k][2];
      count += RectangleUtils.diagonal([
        new XLNT.ExceLintVector(f11.x, f11.y, 0),
        new XLNT.ExceLintVector(f12.x, f12.y, 0),
      ]);
      count += RectangleUtils.diagonal([
        new XLNT.ExceLintVector(f21.x, f21.y, 0),
        new XLNT.ExceLintVector(f22.x, f22.y, 0),
      ]);
    }
    return count;
  }

  public static generate_proposed_fixes(groups: Dict<XLNT.Rectangle[]>): XLNT.ProposedFix[] {
    const proposed_fixes_new = find_all_proposed_fixes(groups);
    // sort by fix metric
    proposed_fixes_new.sort((a, b) => {
      return a.score - b.score;
    });
    return proposed_fixes_new;
  }

  public static merge_groups(groups: Dict<XLNT.Rectangle[]>): Dict<XLNT.Rectangle[]> {
    for (const k of Object.keys(groups)) {
      const g = groups[k].slice();
      groups[k] = this.merge_individual_groups(g);
    }
    return groups;
  }

  public static merge_individual_groups(group: XLNT.Rectangle[]): XLNT.Rectangle[] {
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
            if (!(head_str in deleted_rectangles) && !(working_group_i_str in deleted_rectangles)) {
              updated_rectangles.push(RectangleUtils.bounding_box(head, working_group[i]));
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
        console.log("Too many iterations; abandoning this group.");
        return [[new XLNT.ExceLintVector(-1, -1, 0), new XLNT.ExceLintVector(-1, -1, 0)]];
      }
    }
  }

  // Mark proposed fixes that do not have the same format.
  // Modifies ProposedFix objects, including their scores.
  public static adjust_proposed_fixes(
    fixes: XLNT.ProposedFix[],
    propertiesToGet: XLNT.Spreadsheet,
    origin_col: number,
    origin_row: number
  ): void {
    const proposed_fixes: XLNT.ProposedFix[] = [];
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const fix = fixes[k];
      const rect1 = fix.rect1;
      const rect2 = fix.rect2;

      // Find out which range is "first," i.e., sort by x and then by y.
      const [first, second] =
        XLNT.rectangleComparator(rect1, rect2) <= 0 ? [rect1, rect2] : [rect2, rect1];

      // get the upper-left and bottom-right vectors for the two XLNT.rectangles
      const [ul, _a] = first;
      const [_b, br] = second;

      // get the column and row for the upper-left and bottom-right vectors
      const ul_col = ul.x - origin_col - 1;
      const ul_row = ul.y - origin_row - 1;
      const br_col = br.x - origin_col - 1;
      const br_row = br.y - origin_row - 1;

      // Now check whether the formats are all the same or not.
      // Get the first format and then check that all other cells in the
      // range have the same format.
      // We can iterate over the combination of both ranges at the same
      // time because all proposed fixes must be "merge compatible," i.e.,
      // adjacent XLNT.rectangles that, when merged, form a new rectangle.
      const prop = propertiesToGet[ul_row][ul_col];
      const firstFormat = JSON.stringify(prop);
      for (let i = ul_row; i <= br_row; i++) {
        // if we've already determined that the formats are different
        // stop looking for differences
        if (!fix.sameFormat) {
          break;
        }
        for (let j = ul_col; j <= br_col; j++) {
          const secondFormat = JSON.stringify(propertiesToGet[i][j]);
          if (secondFormat !== firstFormat) {
            // stop looking for differences and modify fix
            fix.sameFormat = false;
            fix.score *= (100 - Config.getFormattingDiscount()) / 100;
            break;
          }
        }
      }
    }
  }
}
