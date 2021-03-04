// Polyfill for flat (IE & Edge)
const flat = require("array.prototype.flat");
flat.shim();

import { ExcelUtils } from "./excelutils";
import { RectangleUtils } from "./rectangleutils";
import { find_all_proposed_fixes } from "./groupme";
import * as XLNT from "./ExceLintTypes";
import { WorkbookOutput } from "./exceljson";
import { Config } from "./config";
import { Classification } from "./classification";
import { Some, None } from "./option";

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
  private static color_list: string[] = [];

  // A hash string indicating no dependencies; in other words,
  // either a formula that makes no references (like `=RAND()`) or a data cell (like `1`)
  private static noDependenciesHash = new XLNT.Fingerprint(12345);

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
  private static filterFixesByUserThreshold(fixes: XLNT.ProposedFix[], thresh: number): XLNT.ProposedFix[] {
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

  /**
   * Performs an incremental ExceLint analysis.
   * @param inp The workbook to analysis.
   * @param prev The previous analysis result, including last observed workbook state.
   * @param edit The sequence of updates to the formula string.
   * @param addr The location of the formula.
   */
  public static update_analysis(
    inp: WorkbookOutput,
    prev: XLNT.WorkbookAnalysis,
    edit: [number, string],
    addr: XLNT.Address
  ): XLNT.WorkbookAnalysis {
    // to make TS compiler stop complaining while I develop
    // TODO REMOVE
    edit = edit;

    // run analysis
    if (false) {
      // incremental analysis
      // TODO
      return prev;
    } else {
      // run the entire big analysis
      return this.process_workbook(inp, addr.worksheet);
    }
  }

  // Given a full analysis, map addresses to rectangles
  public static rectangleDict(a: XLNT.Analysis): XLNT.Dictionary<XLNT.Rectangle> {
    const _d = new XLNT.Dictionary<XLNT.Rectangle>();

    // for every cell in the analysis, find its bounding rectangle
    // and put it in the dictionary, indexed by address (vector)
    const rects = Object.entries(a.grouped_formulas)
      .map(([, value]) => value)
      .flat();
    for (let i = 0; i < rects.length; i++) {
      const rect = rects[i];
      const cells = rect.expand();
      for (let j = 0; j < cells.length; j++) {
        // index this rectangle by this address (vector)
        _d.put(cells[j].asKey(), rect);
      }
    }

    return _d;
  }

  public static refsForRect(r: XLNT.Rectangle): XLNT.ExceLintVector[] {
    r = r;

    return [];
  }

  // Given a set of rectangles indexed by their addresses, produce a set of
  // adjacencies indexed by their addresses
  public static adjacencyDict(rd: XLNT.Dictionary<XLNT.Rectangle>, a: XLNT.Analysis): XLNT.Dictionary<XLNT.Adjacency> {
    const _d = new XLNT.Dictionary<XLNT.Adjacency>();

    // for every cell in the given dictionary, find its adjacencies
    // and index in dictionary by address (vector)
    const addrs = rd.keys;
    for (let i = 0; i < addrs.length; i++) {
      // get the address (a string, because it's a JS dictionary key)
      const addr = addrs[i];

      // get the address vector
      const v = XLNT.ExceLintVector.fromKey(addr);

      // compute the addresses (as keys) of the cells above, below, to the left,
      // and to the right of this cell.
      const up_addr = v.up.asKey();
      const down_addr = v.down.asKey();
      const left_addr = v.left.asKey();
      const right_addr = v.right.asKey();

      // get the rectangle for each adjacency
      // if there is no adjacency (i.e., cell lies on the used range border)
      // store 'None'
      const up_rect = rd.contains(up_addr) ? new Some(rd.get(up_addr)) : None;
      const down_rect = rd.contains(down_addr) ? new Some(rd.get(down_addr)) : None;
      const left_rect = rd.contains(left_addr) ? new Some(rd.get(left_addr)) : None;
      const right_rect = rd.contains(right_addr) ? new Some(rd.get(right_addr)) : None;

      // find fingerprints for adjacent rectangles
      const up_fp = a.formula_fingerprints.get(up_addr);
      const down_fp = a.formula_fingerprints.get(down_addr);
      const left_fp = a.formula_fingerprints.get(left_addr);
      const right_fp = a.formula_fingerprints.get(right_addr);

      // generate tuples
      const up_tup = new XLNT.Tuple2(up_rect, up_fp);
      const down_tup = new XLNT.Tuple2(down_rect, down_fp);
      const left_tup = new XLNT.Tuple2(left_rect, left_fp);
      const right_tup = new XLNT.Tuple2(right_rect, right_fp);

      // add adjacency to dict
      _d.put(addr, new XLNT.Adjacency(up_tup, down_tup, left_tup, right_tup));
    }

    return _d;
  }

  // Prepare to perform an incremental analysis
  public static initIncremental(
    inp: WorkbookOutput,
    sheetName: string,
    beVerbose: boolean = false
  ): XLNT.IncrementalWorkbookAnalysis {
    // look for the requested sheet
    for (let i = 0; i < inp.worksheets.length; i++) {
      const sheet = inp.worksheets[i];

      // skip sheets that don't match sheetName or are empty
      if (Colorize.isNotSameSheet(sheetName, sheet.sheetName) || Colorize.isEmptySheet(sheet)) {
        continue;
      }

      // get the used range
      const usedRangeAddress = Colorize.normalizeAddress(sheet.usedRangeAddress);

      // Get anomalous cells and proposed fixes, among others.
      const a = Colorize.process_suspicious(usedRangeAddress, sheet.formulas, sheet.values, beVerbose);

      // Index rectangles by their component addresses
      const rects = Colorize.rectangleDict(a);

      // Build adjacency map
      const adjs = Colorize.adjacencyDict(rects, a);

      console.log(adjs);
    }

    return new XLNT.IncrementalWorkbookAnalysis();
  }

  // Performs an analysis on an entire workbook
  public static process_workbook(
    inp: WorkbookOutput,
    sheetName: string,
    beVerbose: boolean = false
  ): XLNT.WorkbookAnalysis {
    const wba = new XLNT.WorkbookAnalysis(inp);

    // DEBUG
    let incr = Colorize.initIncremental(inp, sheetName, beVerbose);
    console.log(incr);

    // look for the requested sheet
    for (let i = 0; i < inp.worksheets.length; i++) {
      const sheet = inp.worksheets[i];

      // skip sheets that don't match sheetName or are empty
      if (Colorize.isNotSameSheet(sheetName, sheet.sheetName) || Colorize.isEmptySheet(sheet)) {
        continue;
      }

      // get the used range
      const usedRangeAddress = Colorize.normalizeAddress(sheet.usedRangeAddress);

      // Get anomalous cells and proposed fixes, among others.
      const a = Colorize.process_suspicious(usedRangeAddress, sheet.formulas, sheet.values, beVerbose);

      // Eliminate fixes below user threshold
      a.proposed_fixes = Colorize.filterFixesByUserThreshold(a.proposed_fixes, Config.reportingThreshold);

      // Remove fixes that require fixing both a formula AND formatting.
      // NB: origin_col and origin_row currently hard-coded at 0,0.
      Colorize.adjust_proposed_fixes(a.proposed_fixes, sheet.styles, 0, 0);

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
          if (beVerbose) console.warn("Omitted " + print_formulas + "(too small)");
          continue;
        }

        // Omit fixes with entropy change over threshold
        if (Colorize.fixEntropy(fix) > Config.maxEntropy) {
          const print_formulas = JSON.stringify(rect_info.map((fi) => fi.print_formula));
          if (beVerbose) console.warn("Omitted " + JSON.stringify(print_formulas) + "(too high entropy)");
          continue;
        }

        // Classify fixes & prune based on the best explanation
        const bin = Classification.pruneFixes(Classification.classifyFixes(fix, is_vert, rect_info));

        // IMPORTANT:
        // Exclude reported bugs subject to certain conditions.
        if (Classification.omitFixes(bin, rect_info, beVerbose)) continue;

        // If we're still here, accept this fix
        final_adjusted_fixes.push(fix);

        // Package everything up with the fix
        fix.analysis = new XLNT.FixAnalysis(bin, rect_info, is_vert);
      }

      // gather all statistics about the sheet
      wba.addSheet(new XLNT.WorksheetAnalysis(sheet, final_adjusted_fixes, a));
    }
    return wba;
  }

  /**
   * Find all the fingerprints for all the formulas in the given used range.
   * This is the actual fingerprint implementation; fingerprintFormulas is a
   * helper method.
   * @param formulas A spreadsheet of formula strings.
   * @param origin_col A column offset from which to start.
   * @param origin_row A row offset from which to start.
   */
  private static fingerprintFormulasImpl(
    formulas: XLNT.Spreadsheet,
    origin_col: number,
    origin_row: number
  ): XLNT.Dictionary<XLNT.Fingerprint> {
    const base_vector = ExcelUtils.baseVector();
    const _d = new XLNT.Dictionary<XLNT.Fingerprint>();

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
              const v = new XLNT.ExceLintVector(adjustedX, adjustedY, 0);
              _d.put(v.asKey(), Colorize.noDependenciesHash);
            }
          } else {
            // compute resultant vector
            const vec = vec_array.reduce(XLNT.ExceLintVector.VectorSum);

            // get the address vector
            const v = new XLNT.ExceLintVector(adjustedX, adjustedY, 0);

            // add to dict
            if (vec.equals(base_vector)) {
              _d.put(v.asKey(), Colorize.noDependenciesHash);
            } else {
              const hash = vec.hash();
              _d.put(v.asKey(), new XLNT.Fingerprint(hash));
            }
          }
        }
      }
    }
    return _d;
  }

  // Returns all referenced data so it can be colored later.
  public static color_all_data(refs: XLNT.Dictionary<boolean>): XLNT.Dictionary<XLNT.Fingerprint> {
    const _d = new XLNT.Dictionary<XLNT.Fingerprint>();

    for (const refvec of refs.keys) {
      _d.put(refvec, Colorize.noDependenciesHash);
    }

    return _d;
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
            value_array.push([new XLNT.ExceLintVector(adjustedX, adjustedY, 1), Colorize.noDependenciesHash]);
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
    data_fingerprints: XLNT.Dictionary<XLNT.Fingerprint>,
    sortfn: (n1: XLNT.ExceLintVector, n2: XLNT.ExceLintVector) => number
  ): XLNT.Dictionary<XLNT.ExceLintVector[]> {
    // Separate into groups based on their XLNT.Fingerprint value.
    const groups = new XLNT.Dictionary<XLNT.ExceLintVector[]>();
    for (const key of data_fingerprints.keys) {
      const vec = XLNT.ExceLintVector.fromKey(key);
      const fp = data_fingerprints.get(key).asKey();
      if (!groups.contains(fp)) {
        groups.put(fp, []); // initialize array if necessary
      }
      groups.get(fp).push(vec);
    }
    // Now sort them all.
    for (const z of groups.keys) {
      groups.get(z).sort(sortfn);
    }
    return groups;
  }

  // Collect all ranges of cells that share a XLNT.Fingerprint
  private static find_contiguous_regions(
    groups: XLNT.Dictionary<XLNT.ExceLintVector[]>
  ): XLNT.Dictionary<XLNT.Rectangle[]> {
    const output = new XLNT.Dictionary<XLNT.Rectangle[]>();

    for (const key of groups.keys) {
      // Here, we scan all of the vectors in this group, accumulating
      // all adjacent vectors by tracking the start and end. Whevener
      // we encounter a non-adjacent vector, push the region to the output
      // list and then start tracking a new region.
      output.put(key, []); // initialize
      let start = groups.get(key).shift() as XLNT.ExceLintVector; // remove the first vector from the list
      let end = start;
      for (const v of groups.get(key)) {
        // Check if v is in the same column as last, adjacent row
        if (v.x === end.x && v.y === end.y + 1) {
          end = v;
        } else {
          output.get(key).push(new XLNT.Rectangle(start, end));
          start = v;
          end = v;
        }
      }
      output.get(key).push(new XLNT.Rectangle(start, end));
    }
    return output;
  }

  public static identify_groups(
    data_fingerprints: XLNT.Dictionary<XLNT.Fingerprint>
  ): XLNT.Dictionary<XLNT.Rectangle[]> {
    const id = Colorize.identify_ranges(data_fingerprints, ExcelUtils.ColumnSort);
    const gr = Colorize.find_contiguous_regions(id);
    // Now try to merge stuff with the same hash.
    const newGr1 = gr.clone();
    const mg = Colorize.merge_groups(newGr1);
    return mg;
  }

  // public static generate_suspicious_cells(
  //   cols: number,
  //   rows: number,
  //   origin_col: number,
  //   origin_row: number,
  //   matrix: Array<Array<number>>,
  //   probs: Array<Array<number>>,
  //   threshold = 0.01
  // ): Array<XLNT.ExceLintVector> {
  //   const cells = [];
  //   let sumValues = 0;
  //   let countValues = 0;
  //   for (let i = 0; i < cols; i++) {
  //     for (let j = 0; j < rows; j++) {
  //       const adjustedX = j + origin_col + 1;
  //       const adjustedY = i + origin_row + 1;
  //       if (probs[i][j] > 0) {
  //         sumValues += matrix[i][j];
  //         countValues += 1;
  //         if (probs[i][j] <= threshold) {
  //           if (matrix[i][j] !== 0) {
  //             // Never push an empty cell.
  //             cells.push([adjustedX, adjustedY, probs[i][j]]);
  //           }
  //         }
  //       }
  //     }
  //   }
  //   const avgValues = sumValues / countValues;
  //   cells.sort((a, b) => {
  //     return Math.abs(b[2] - avgValues) - Math.abs(a[2] - avgValues);
  //   });
  //   return cells;
  // }

  /**
   * Determine whether the number of formulas in the spreadsheet exceeds
   * a hand-tuned threshold (for analysis responsiveness).
   * @param formulas A Spreadsheet of formulas
   */
  public static tooManyFormulas(formulas: XLNT.Spreadsheet) {
    const totalFormulas = (formulas as any).flat().filter(Boolean).length;
    return totalFormulas > Config.formulasThreshold;
  }

  /**
   * Determine whether the number of values in the spreadsheet exceeds
   * a hand-tuned threshold (for analysis responsiveness).
   * @param values A Spreadsheet of values
   */
  public static tooManyValues(values: XLNT.Spreadsheet) {
    const totalValues = (values as any).flat().filter(Boolean).length;
    return totalValues > Config.valuesThreshold;
  }

  /**
   * Find all the fingerprints for all the formulas in the given used range.
   * TODO FIX: I'm not exactly sure how the used range is used here.
   * @param usedRangeAddress A1 string representation of used range reference
   * @param formulas A spreadsheet of formulas.
   * @param beVerbose Print diagnostics to console when true.
   */
  public static fingerprintFormulas(
    usedRangeAddress: string,
    formulas: XLNT.Spreadsheet,
    beVerbose: boolean
  ): XLNT.Dictionary<XLNT.Fingerprint> {
    const [, startCell] = ExcelUtils.extract_sheet_cell(usedRangeAddress);
    const origin = ExcelUtils.cell_dependency(startCell, 0, 0);

    // Filter out non-empty items from whole matrix.
    if (Colorize.tooManyFormulas(formulas)) {
      if (beVerbose) console.warn("Too many formulas to perform formula analysis.");
      return new XLNT.Dictionary<XLNT.Fingerprint>();
    } else {
      return Colorize.fingerprintFormulasImpl(formulas, origin.x - 1, origin.y - 1);
    }
  }

  /**
   * Find all the fingerprints for all the data in the given used range.
   * TODO FIX: I'm not exactly sure how the used range is used here.
   * @param usedRangeAddress A1 string representation of used range reference
   * @param formulas A spreadsheet of formula strings.
   * @param values A spreadsheet of values.
   * @param beVerbose Print diagnostics to console when true.
   */
  public static fingerprintData(
    usedRangeAddress: string,
    formulas: XLNT.Spreadsheet,
    values: XLNT.Spreadsheet,
    beVerbose: boolean
  ): XLNT.Dictionary<XLNT.Fingerprint> {
    const [, startCell] = ExcelUtils.extract_sheet_cell(usedRangeAddress);
    const origin = ExcelUtils.cell_dependency(startCell, 0, 0);

    // Filter out non-empty items from whole matrix.
    if (Colorize.tooManyValues(values)) {
      if (beVerbose) console.warn("Too many values to perform reference analysis.");
      return new XLNT.Dictionary<XLNT.Fingerprint>();
    } else {
      // Compute references (to color referenced data).
      const refs: XLNT.Dictionary<boolean> = ExcelUtils.generate_all_references(formulas, origin.x - 1, origin.y - 1);

      return Colorize.color_all_data(refs);
    }
  }

  /**
   * This is the core of an ExceLint analysis.  It fingerprints formulas and data,
   * partitions them into rectangular regions, and then returns an Analysis object
   * that contains ProposedFixes.
   * @param usedRangeAddr A1 string representation of used range reference.
   * @param formulas A spreadsheet of formula strings.
   * @param values A spreadsheet of value (data) strings.
   * @param beVerbose Print diagnostics to console when true.
   */
  public static process_suspicious(
    usedRangeAddr: string,
    formulas: XLNT.Spreadsheet,
    values: XLNT.Spreadsheet,
    beVerbose: boolean
  ): XLNT.Analysis {
    // fingerprint all the formulas
    const processed_formulas = this.fingerprintFormulas(usedRangeAddr, formulas, beVerbose);

    // fingerprint all the data
    const referenced_data = Colorize.fingerprintData(usedRangeAddr, formulas, values, beVerbose);

    // find regions for data
    const grouped_data = Colorize.identify_groups(referenced_data);

    // find regions for formulas
    const grouped_formulas = Colorize.identify_groups(processed_formulas);

    // Identify suspicious cells (disabled)
    let suspicious_cells: XLNT.ExceLintVector[] = [];

    // find proposed fixes
    const proposed_fixes = Colorize.generate_proposed_fixes(grouped_formulas);

    return new XLNT.Analysis(
      suspicious_cells,
      grouped_formulas,
      grouped_data,
      proposed_fixes,
      processed_formulas,
      referenced_data
    );
  }

  // Shannon entropy.
  public static entropy(p: number): number {
    return -p * Math.log2(p);
  }

  // Take two counts and compute the normalized entropy difference that would result if these were 'merged'.
  public static entropydiff(oldcount1: number, oldcount2: number) {
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
    const t1 = target.upperleft;
    const t2 = target.bottomright;
    const m1 = merge_with.upperleft;
    const m2 = merge_with.bottomright;
    const n_target = RectangleUtils.area(
      new XLNT.Rectangle(new XLNT.ExceLintVector(t1.x, t1.y, 0), new XLNT.ExceLintVector(t2.x, t2.y, 0))
    );
    const n_merge_with = RectangleUtils.area(
      new XLNT.Rectangle(new XLNT.ExceLintVector(m1.x, m1.y, 0), new XLNT.ExceLintVector(m2.x, m2.y, 0))
    );
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
  public static count_proposed_fixes(fixes: Array<[number, XLNT.Rectangle, XLNT.Rectangle]>): number {
    let count = 0;
    // tslint:disable-next-line: forin
    for (const k in fixes) {
      const f11 = fixes[k][1].upperleft;
      const f12 = fixes[k][1].bottomright;
      const f21 = fixes[k][2].upperleft;
      const f22 = fixes[k][2].bottomright;
      count += RectangleUtils.diagonal(
        new XLNT.Rectangle(new XLNT.ExceLintVector(f11.x, f11.y, 0), new XLNT.ExceLintVector(f12.x, f12.y, 0))
      );
      count += RectangleUtils.diagonal(
        new XLNT.Rectangle(new XLNT.ExceLintVector(f21.x, f21.y, 0), new XLNT.ExceLintVector(f22.x, f22.y, 0))
      );
    }
    return count;
  }

  public static generate_proposed_fixes(groups: XLNT.Dictionary<XLNT.Rectangle[]>): XLNT.ProposedFix[] {
    const proposed_fixes_new = find_all_proposed_fixes(groups);
    // sort by fix metric
    proposed_fixes_new.sort((a, b) => {
      return a.score - b.score;
    });
    return proposed_fixes_new;
  }

  public static merge_groups(groups: XLNT.Dictionary<XLNT.Rectangle[]>): XLNT.Dictionary<XLNT.Rectangle[]> {
    for (const k of groups.keys) {
      const g = groups.get(k).slice();
      groups.put(k, Colorize.merge_individual_groups(g));
    }
    return groups;
  }

  public static merge_individual_groups(group: XLNT.Rectangle[]): XLNT.Rectangle[] {
    let numIterations = 0;
    group = group.sort();
    while (true) {
      let merged_one = false;
      const deleted_rectangles = new XLNT.Dictionary<boolean>();
      const updated_rectangles = [];
      const working_group = group.slice();
      while (working_group.length > 0) {
        const head = working_group.shift() as XLNT.Rectangle;
        for (let i = 0; i < working_group.length; i++) {
          if (RectangleUtils.is_mergeable(head, working_group[i])) {
            const head_str = JSON.stringify(head);
            const working_group_i_str = JSON.stringify(working_group[i]);
            // NB: 12/7/19 New check below, used to be unconditional.
            if (!(head_str in deleted_rectangles) && !(working_group_i_str in deleted_rectangles)) {
              updated_rectangles.push(RectangleUtils.bounding_box(head, working_group[i]));
              deleted_rectangles.put(head_str, true);
              deleted_rectangles.put(working_group_i_str, true);
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
        const tl = new XLNT.ExceLintVector(-1, -1, 0);
        const br = new XLNT.ExceLintVector(-1, -1, 0);
        return [new XLNT.Rectangle(tl, br)];
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
    for (const k in fixes) {
      const fix = fixes[k];
      const rect1 = fix.rect1;
      const rect2 = fix.rect2;

      // Find out which range is "first," i.e., sort by x and then by y.
      const [first, second] = XLNT.rectangleComparator(rect1, rect2) <= 0 ? [rect1, rect2] : [rect2, rect1];

      // get the upper-left and bottom-right vectors for the two XLNT.rectangles
      const ul = first.upperleft;
      const br = second.bottomright;

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
