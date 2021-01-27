// Polyfill for flat (IE & Edge)
const flat = require("array.prototype.flat");
flat.shim();

import { ExcelUtils } from "./excelutils";
import { RectangleUtils } from "./rectangleutils";
import { Timer } from "./timer";
import { JSONclone } from "./jsonclone";
import { find_all_proposed_fixes } from "./groupme";
import { Stencil } from "./infogain";
import {
  ExceLintVector,
  Dict,
  Spreadsheet,
  Fingerprint,
  Rectangle,
  ProposedFix,
  Metric,
  Analysis,
  rectangleComparator,
  rectangles,
  rect1,
  rect2,
  upperleft,
  bottomright,
  expand,
  WorksheetAnalysis,
  WorkbookAnalysis,
  FixAnalysis,
  RectInfo,
} from "./ExceLintTypes";
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
  private static filterFixesByUserThreshold(fixes: ProposedFix[], thresh: number): ProposedFix[] {
    const fixes2: ProposedFix[] = [];
    for (let ind = 0; ind < fixes.length; ind++) {
      const [score, first, second] = fixes[ind];
      let adjusted_score = -score;
      if (adjusted_score * 100 >= thresh) {
        fixes2.push([adjusted_score, first, second]);
      }
    }
    return fixes2;
  }

  // Returns true if the "direction" of a fix is vertical
  private static fixIsVertical(fix: ProposedFix): boolean {
    const rect1_ul_x = upperleft(rect1(fix)).x;
    const rect2_ul_x = upperleft(rect2(fix)).x;
    return rect1_ul_x === rect2_ul_x;
  }

  private static fixCellCount(fix: ProposedFix): number {
    const fixRange = expand(upperleft(rect1(fix)), bottomright(rect1(fix))).concat(
      expand(upperleft(rect2(fix)), bottomright(rect2(fix)))
    );
    return fixRange.length;
  }

  private static fixEntropy(fix: ProposedFix): number {
    const leftFixSize = expand(upperleft(rect1(fix)), bottomright(rect1(fix))).length;
    const rightFixSize = expand(upperleft(rect2(fix)), bottomright(rect2(fix))).length;
    const totalSize = leftFixSize + rightFixSize;
    const fixEntropy = -(
      (leftFixSize / totalSize) * Math.log2(leftFixSize / totalSize) +
      (rightFixSize / totalSize) * Math.log2(rightFixSize / totalSize)
    );
    return fixEntropy;
  }

  // Performs an analysis on an entire workbook
  public static process_workbook(inp: WorkbookOutput, sheetName: string): WorkbookAnalysis {
    const wba = new WorkbookAnalysis();

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
      const final_adjusted_fixes: ProposedFix[] = []; // We will eventually trim these.
      a.proposed_fixes = Colorize.filterFixesByUserThreshold(
        a.proposed_fixes,
        Config.reportingThreshold
      );

      // Remove fixes that require fixing both a formula AND formatting.
      // NB: origin_col and origin_row currently hard-coded at 0,0.
      const initial_adjusted_fixes = Colorize.adjust_proposed_fixes(
        a.proposed_fixes,
        sheet.styles,
        0,
        0
      );

      // Process all the fixes, classifying and optionally pruning them.
      const example_fixes_r1c1: FixAnalysis[] = []; // TODO DAN: this really needs a type

      for (let ind = 0; ind < initial_adjusted_fixes.length; ind++) {
        // Get this fix
        const fix = initial_adjusted_fixes[ind];

        // Determine the direction of the range (vertical or horizontal) by looking at the axes.
        const is_vert: boolean = Colorize.fixIsVertical(fix);

        // Formula info for each rectangle
        const rect_info = rectangles(fix).map((rect) => new RectInfo(rect, sheet));

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

        // Package everything up and append to array
        example_fixes_r1c1.push(new FixAnalysis(fix, bin, rect_info, is_vert));
      }

      let elapsed = myTimer.elapsedTime();
      if (Config.noElapsedTime) {
        elapsed = 0; // Dummy value, used for regression testing.
      }

      // gather all statistics about the sheet
      wba.appendSheet(new WorksheetAnalysis(sheet, final_adjusted_fixes));
    }
    return wba;
  }

  // Generate dependence vectors and their hash for all formulas.
  public static process_formulas(
    formulas: Spreadsheet,
    origin_col: number,
    origin_row: number
  ): [ExceLintVector, Fingerprint][] {
    const base_vector = ExcelUtils.baseVector();
    const output: Array<[ExceLintVector, Fingerprint]> = [];

    // Compute the vectors for all of the formulas.
    for (let i = 0; i < formulas.length; i++) {
      const row = formulas[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        // If it's a formula, process it.
        if (cell.length > 0) {
          // FIXME MAYBE  && (row[j][0] === '=')) {
          const vec_array: ExceLintVector[] = ExcelUtils.all_dependencies(
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
                new ExceLintVector(adjustedX, adjustedY, 0),
                Colorize.noDependenciesHash,
              ]);
            }
          } else {
            const vec = vec_array.reduce(ExceLintVector.VectorSum);
            if (vec.equals(base_vector)) {
              // No dependencies! Use a distinguished value.
              // Emery's FIXME: RESTORE THIS output.push([[adjustedX, adjustedY, 0], Colorize.distinguishedZeroHash]);
              // DAN TODO: I don't understand this case.
              output.push([
                new ExceLintVector(adjustedX, adjustedY, 0),
                Colorize.noDependenciesHash,
              ]);
            } else {
              const hash = vec.hash();
              output.push([new ExceLintVector(adjustedX, adjustedY, 0), hash.toString()]);
            }
          }
        }
      }
    }
    return output;
  }

  // Returns all referenced data so it can be colored later.
  public static color_all_data(refs: Dict<boolean>): [ExceLintVector, Fingerprint][] {
    const referenced_data: [ExceLintVector, Fingerprint][] = [];
    for (const refvec of Object.keys(refs)) {
      const rv = refvec.split(",");
      const row = Number(rv[0]);
      const col = Number(rv[1]);
      referenced_data.push([new ExceLintVector(row, col, 0), Colorize.noDependenciesHash]); // See comment at top of function declaration.
    }
    return referenced_data;
  }

  // Take all values and return an array of each row and column.
  // Note that for now, the last value of each tuple is set to 1.
  public static process_values(
    values: Spreadsheet,
    formulas: Spreadsheet,
    origin_col: number,
    origin_row: number
  ): [ExceLintVector, Fingerprint][] {
    const value_array: [ExceLintVector, Fingerprint][] = [];
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
              new ExceLintVector(adjustedX, adjustedY, 1),
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
    list: Array<[ExceLintVector, string]>,
    sortfn: (n1: ExceLintVector, n2: ExceLintVector) => number
  ): Dict<ExceLintVector[]> {
    // Separate into groups based on their fingerprint value.
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

  // Collect all ranges of cells that share a fingerprint
  private static find_contiguous_regions(groups: Dict<ExceLintVector[]>): Dict<Rectangle[]> {
    const output: Dict<Rectangle[]> = {};

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

  public static identify_groups(theList: [ExceLintVector, string][]): Dict<Rectangle[]> {
    const id = Colorize.identify_ranges(theList, ExcelUtils.ColumnSort);
    const gr = Colorize.find_contiguous_regions(id);
    // Now try to merge stuff with the same hash.
    const newGr1 = JSONclone.clone(gr);
    const mg = Colorize.merge_groups(newGr1);
    return mg;
  }

  public static processed_to_matrix(
    cols: number,
    rows: number,
    origin_col: number,
    origin_row: number,
    processed: Array<[ExceLintVector, string]>
  ): Array<Array<number>> {
    // Invert the hash table.
    // First, initialize a zero-filled matrix.
    const matrix: Array<Array<number>> = new Array(cols);
    for (let i = 0; i < cols; i++) {
      matrix[i] = new Array(rows).fill(0);
    }
    // Now iterate through the processed formulas and update the matrix.
    for (const item of processed) {
      const [vect, val] = item;
      // Yes, I know this is confusing. Will fix later.
      //	    console.log('C) cols = ' + rows + ', rows = ' + cols + '; row = ' + row + ', col = ' + col);
      const adjustedX = vect.y - origin_row - 1;
      const adjustedY = vect.x - origin_col - 1;
      let value = Number(Colorize.noDependenciesHash);
      if (vect.isConstant()) {
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
  ): Array<ExceLintVector> {
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
    formulas: Spreadsheet,
    values: Spreadsheet
  ): Analysis {
    const t = new Timer("process_suspicious");

    const [sheetName, startCell] = ExcelUtils.extract_sheet_cell(usedRangeAddress);
    const origin = ExcelUtils.cell_dependency(startCell, 0, 0);

    let processed_formulas: [ExceLintVector, string][] = [];
    // Filter out non-empty items from whole matrix.
    const totalFormulas = (formulas as any).flat().filter(Boolean).length;

    if (totalFormulas > Config.formulasThreshold) {
      console.warn("Too many formulas to perform formula analysis.");
    } else {
      processed_formulas = Colorize.process_formulas(formulas, origin.x - 1, origin.y - 1);
    }

    let referenced_data: [ExceLintVector, Fingerprint][] = [];
    let data_values: [ExceLintVector, Fingerprint][] = [];
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
    let suspicious_cells: ExceLintVector[] = [];

    // find proposed fixes
    const proposed_fixes = Colorize.generate_proposed_fixes(grouped_formulas);

    return new Analysis(suspicious_cells, grouped_formulas, grouped_data, proposed_fixes);
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
    target: Rectangle,
    merge_with_norm: number,
    merge_with: Rectangle
  ): Metric {
    //	console.log('fix_metric: ' + target_norm + ', ' + JSON.stringify(target) + ', ' + merge_with_norm + ', ' + JSON.stringify(merge_with));
    const [t1, t2] = target;
    const [m1, m2] = merge_with;
    const n_target = RectangleUtils.area([
      new ExceLintVector(t1.x, t1.y, 0),
      new ExceLintVector(t2.x, t2.y, 0),
    ]);
    const n_merge_with = RectangleUtils.area([
      new ExceLintVector(m1.x, m1.y, 0),
      new ExceLintVector(m2.x, m2.y, 0),
    ]);
    const n_min = Math.min(n_target, n_merge_with);
    const n_max = Math.max(n_target, n_merge_with);
    const norm_min = Math.min(merge_with_norm, target_norm);
    const norm_max = Math.max(merge_with_norm, target_norm);
    let fix_distance = Math.abs(norm_max - norm_min) / ExceLintVector.Multiplier;

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
    fixes: Array<[number, [ExceLintVector, ExceLintVector], [ExceLintVector, ExceLintVector]]>
  ): number {
    let count = 0;
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const [f11, f12] = fixes[k][1];
      const [f21, f22] = fixes[k][2];
      count += RectangleUtils.diagonal([
        new ExceLintVector(f11.x, f11.y, 0),
        new ExceLintVector(f12.x, f12.y, 0),
      ]);
      count += RectangleUtils.diagonal([
        new ExceLintVector(f21.x, f21.y, 0),
        new ExceLintVector(f22.x, f22.y, 0),
      ]);
    }
    return count;
  }

  // Try to merge fixes into larger groups.
  public static fix_proposed_fixes(
    fixes: Array<[number, [ExceLintVector, ExceLintVector], [ExceLintVector, ExceLintVector]]>
  ): Array<[number, [ExceLintVector, ExceLintVector], [ExceLintVector, ExceLintVector]]> {
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
      if (-original_score < Config.reportingThreshold / 100) {
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
          const newscore = -original_score * JSON.parse(back[this_front_str][0]);
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
          const newscore = -original_score * JSON.parse(front[this_back_str][0]);
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

  public static generate_proposed_fixes(groups: Dict<Rectangle[]>): ProposedFix[] {
    const proposed_fixes_new = find_all_proposed_fixes(groups);
    // sort by fix metric
    proposed_fixes_new.sort((a, b) => {
      return a[0] - b[0];
    });
    return proposed_fixes_new;
  }

  public static merge_groups(groups: Dict<Rectangle[]>): Dict<Rectangle[]> {
    for (const k of Object.keys(groups)) {
      const g = groups[k].slice();
      groups[k] = this.merge_individual_groups(g);
    }
    return groups;
  }

  public static merge_individual_groups(group: Rectangle[]): Rectangle[] {
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
        return [[new ExceLintVector(-1, -1, 0), new ExceLintVector(-1, -1, 0)]];
      }
    }
  }

  // Filter out any proposed fixes that do not have the same format.
  public static adjust_proposed_fixes(
    fixes: ProposedFix[],
    propertiesToGet: Spreadsheet,
    origin_col: number,
    origin_row: number
  ): ProposedFix[] {
    const proposed_fixes: ProposedFix[] = [];
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      // Format of proposed fixes =, e.g., [-3.016844756293869, [[5,7],[5,11]],[[6,7],[6,11]]]
      // entropy, and two ranges:
      //    upper-left corner of range (column, row), lower-right corner of range (column, row)

      const [score, rect1, rect2] = fixes[k];

      // Find out which range is "first," i.e., sort by x and then by y.
      const [first, second] =
        rectangleComparator(rect1, rect2) <= 0 ? [rect1, rect2] : [rect2, rect1];

      // get the upper-left and bottom-right vectors for the two rectangles
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
      // adjacent rectangles that, when merged, form a new rectangle.
      let sameFormats = true;
      const prop = propertiesToGet[ul_row][ul_col];
      const firstFormat = JSON.stringify(prop);
      for (let i = ul_row; i <= br_row; i++) {
        for (let j = ul_col; j <= br_col; j++) {
          const str = JSON.stringify(propertiesToGet[i][j]);
          if (str !== firstFormat) {
            sameFormats = false;
            break;
          }
        }
      }
      proposed_fixes.push([score, first, second]);
    }
    return proposed_fixes;
  }

  public static find_suspicious_cells(
    cols: number,
    rows: number,
    origin: ExceLintVector,
    formulas: any[][],
    processed_formulas: [ExceLintVector, string][],
    data_values: [ExceLintVector, string][],
    threshold: number
  ): ExceLintVector[] {
    return []; // FIXME disabled for now
    let suspiciousCells: any[];
    {
      // data_values = data_values;
      const formula_matrix = Colorize.processed_to_matrix(
        cols,
        rows,
        origin.x - 1,
        origin.y - 1,
        processed_formulas.concat(data_values)
      );

      const stencil = Colorize.stencilize(formula_matrix);
      // console.log('after stencilize: stencil = ' + JSON.stringify(stencil));
      const probs = Colorize.compute_stencil_probabilities(cols, rows, stencil);
      // console.log('probs = ' + JSON.stringify(probs));

      const candidateSuspiciousCells = Colorize.generate_suspicious_cells(
        cols,
        rows,
        origin.x - 1,
        origin.y - 1,
        formula_matrix,
        probs,
        threshold
      );
      // Prune any cell that is in fact a formula.

      if (typeof formulas !== "undefined") {
        let totalFormulaWeight = 0;
        let totalWeight = 0;
        suspiciousCells = candidateSuspiciousCells.filter((c) => {
          const theFormula = formulas[c[1] - origin[1]][c[0] - origin[0]];
          if (theFormula.length < 1 || theFormula[0] !== "=") {
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
        console.log("after processing 1, suspiciousCells = " + JSON.stringify(suspiciousCells));
        suspiciousCells = suspiciousCells.map((c) => [c[0], c[1], c[2] * multiplier]);
        console.log("after processing 2, suspiciousCells = " + JSON.stringify(suspiciousCells));
        suspiciousCells = suspiciousCells.filter((c) => c[2] <= threshold);
        console.log("after processing 3, suspiciousCells = " + JSON.stringify(suspiciousCells));
      } else {
        suspiciousCells = candidateSuspiciousCells;
      }
    }
    return suspiciousCells;
  }
}
