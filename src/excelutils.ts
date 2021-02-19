// excel-utils
// Emery Berger, Microsoft Research / UMass Amherst
// https://emeryberger.com

import * as sjcl from "sjcl";
import { RectangleUtils } from "./rectangleutils";
import { ProposedFix, ExceLintVector, Dict, Spreadsheet, upperleft, bottomright, Address } from "./ExceLintTypes";

export class ExcelUtils {
  // sort routine
  static readonly ColumnSort = (a: ExceLintVector, b: ExceLintVector) => {
    if (a.x === b.x) {
      return a.y - b.y;
    } else {
      return a.x - b.x;
    }
  };

  // Matchers for all kinds of Excel expressions.
  private static general_re = "\\$?[A-Z][A-Z]?\\$?[\\d\\u2000-\\u6000]+"; // column and row number, optionally with $
  private static sheet_re = "[^\\!]+";
  private static sheet_plus_cell = new RegExp("(" + ExcelUtils.sheet_re + ")\\!(" + ExcelUtils.general_re + ")");
  private static sheet_plus_range = new RegExp(
    "(" + ExcelUtils.sheet_re + ")\\!(" + ExcelUtils.general_re + "):(" + ExcelUtils.general_re + ")"
  );
  public static single_dep = new RegExp("(" + ExcelUtils.general_re + ")");
  public static range_pair = new RegExp("(" + ExcelUtils.general_re + "):(" + ExcelUtils.general_re + ")", "g");
  private static number_dep = new RegExp("([0-9]+\\.?[0-9]*)");
  public static cell_both_relative = new RegExp("[^\\$A-Z]?([A-Z][A-Z]?)([\\d\\u2000-\\u6000]+)");
  public static cell_col_absolute = new RegExp("\\$([A-Z][A-Z]?)([\\d\\u2000-\\u6000]+)");
  public static cell_row_absolute = new RegExp("[^\\$A-Z]?([A-Z][A-Z]?)\\$([\\d\\u2000-\\u6000]+)");
  public static cell_both_absolute = new RegExp("\\$([A-Z][A-Z]?)\\$([\\d\\u2000-\\u6000]+)");

  // We need to filter out all formulas with these characteristics so they don't mess with our dependency regexps.

  private static formulas_with_numbers = new RegExp(
    "/ATAN2|BIN2DEC|BIN2HEX|BIN2OCT|DAYS360|DEC2BIN|DEC2HEX|DEC2OCT|HEX2BIN|HEX2DEC|HEX2OCT|IMLOG2|IMLOG10|LOG10|OCT2BIN|OCT2DEC|OCT2HEX|SUNX2MY2|SUMX2PY2|SUMXMY2|T.DIST.2T|T.INV.2T/",
    "g"
  );
  // Same with sheet name references.
  private static formulas_with_quoted_sheetnames_1 = new RegExp("'[^']*'!" + "\\$?[A-Z][A-Z]?\\$?\\d+", "g");
  private static formulas_with_quoted_sheetnames_2 = new RegExp(
    "'[^']*'!" + "\\$?[A-Z][A-Z]?\\$?\\d+" + ":" + "\\$?[A-Z][A-Z]?\\$?\\d+",
    "g"
  );
  private static formulas_with_unquoted_sheetnames_1 = new RegExp("[A-Za-z0-9]+!" + "\\$?[A-Z][A-Z]?\\$?\\d+", "g");
  private static formulas_with_unquoted_sheetnames_2 = new RegExp(
    "[A-Za-z0-9]+!" + "\\$?[A-Z][A-Z]?\\$?\\d+" + ":" + "\\$?[A-Z][A-Z]?\\$?\\d+",
    "g"
  );
  private static formulas_with_structured_references = new RegExp("\\[([^\\]])*\\]", "g");

  private static originalSheetSuffix = "_EL";

  // Get the saved formats for this sheet (by its unique identifier).
  public static saved_original_sheetname(id: string): string {
    return this.hash_sheet(id, 28) + this.originalSheetSuffix;
  }

  // Convert the UID string into a hashed version using SHA256, truncated to a max length.
  public static hash_sheet(uid: string, maxlen = 31): string {
    // We can't just use the UID because it is too long to be a sheet name in Excel (limit is 31 characters).
    return sjcl.codec.base32.fromBits(sjcl.hash.sha256.hash(uid)).slice(0, maxlen);
  }

  public static get_rectangle(proposed_fixes: ProposedFix[], current_fix: number): [string, string, string, string] {
    if (!proposed_fixes) {
      return null;
    }
    if (proposed_fixes.length > 0) {
      const r = RectangleUtils.bounding_box(proposed_fixes[current_fix].rect1, proposed_fixes[current_fix].rect2);
      // convert to sheet notation
      const col0 = ExcelUtils.column_index_to_name(upperleft(r).x);
      const row0 = upperleft(r).y.toString();
      const col1 = ExcelUtils.column_index_to_name(bottomright(r).x);
      const row1 = bottomright(r).y.toString();
      return [col0, row0, col1, row1];
    } else {
      return null;
    }
  }

  // Take a range string and compute the number of cells.
  public static get_number_of_cells(address: string): number {
    // Compute the number of cells in the range "usedRange".
    const usedRangeAddresses = ExcelUtils.extract_sheet_range(address);
    const upperLeftCorner = ExcelUtils.cell_dependency(usedRangeAddresses[1], 0, 0);
    const lowerRightCorner = ExcelUtils.cell_dependency(usedRangeAddresses[2], 0, 0);
    const numberOfCellsUsed = RectangleUtils.area([upperLeftCorner, lowerRightCorner]);
    return numberOfCellsUsed;
  }

  // Convert an Excel column name (a string of alphabetical charcaters) into a number.
  public static column_name_to_index(name: string): number {
    if (name.length === 1) {
      // optimizing for the overwhelmingly common case
      return name[0].charCodeAt(0) - "A".charCodeAt(0) + 1;
    }
    let value = 0;
    const split_name = name.split("");
    for (const i of split_name) {
      value *= 26;
      value += i.charCodeAt(0) - "A".charCodeAt(0) + 1;
    }
    return value;
  }

  // Convert a column number to a name (as in, 3 => 'C').
  public static column_index_to_name(index: number): string {
    let str = "";
    while (index > 0) {
      str += String.fromCharCode(((index - 1) % 26) + 65); // 65 = 'A'
      index = Math.floor((index - 1) / 26);
    }
    return str.split("").reverse().join("");
  }

  // Returns a vector (x, y) corresponding to the column and row of the computed dependency.
  public static cell_dependency(cell: string, origin_col: number, origin_row: number): ExceLintVector {
    const alwaysReturnAdjustedColRow = false;
    {
      const r = ExcelUtils.cell_both_absolute.exec(cell);
      if (r) {
        const col = ExcelUtils.column_name_to_index(r[1]);
        let row = Number(r[2]);
        if (r[2][0] >= "\u2000") {
          row = Number(r[2].charCodeAt(0) - 16384);
        }
        if (alwaysReturnAdjustedColRow) {
          return new ExceLintVector(col - origin_col, row - origin_row, 0);
        } else {
          return new ExceLintVector(col, row, 0);
        }
      }
    }

    {
      const r = ExcelUtils.cell_col_absolute.exec(cell);
      if (r) {
        const col = ExcelUtils.column_name_to_index(r[1]);
        let row = Number(r[2]);
        if (r[2][0] >= "\u2000") {
          row = Number(r[2].charCodeAt(0) - 16384);
        }
        if (alwaysReturnAdjustedColRow) {
          return new ExceLintVector(col, row, 0);
        } else {
          return new ExceLintVector(col, row - origin_row, 0);
        }
      }
    }

    {
      const r = ExcelUtils.cell_row_absolute.exec(cell);
      if (r) {
        const col = ExcelUtils.column_name_to_index(r[1]);
        let row = Number(r[2]);
        if (r[2][0] >= "\u2000") {
          row = Number(r[2].charCodeAt(0) - 16384);
        }
        if (alwaysReturnAdjustedColRow) {
          return new ExceLintVector(col, row, 0);
        } else {
          return new ExceLintVector(col - origin_col, row, 0);
        }
      }
    }

    {
      const r = ExcelUtils.cell_both_relative.exec(cell);
      if (r) {
        const col = ExcelUtils.column_name_to_index(r[1]);
        let row = Number(r[2]);
        if (r[2][0] >= "\u2000") {
          row = Number(r[2].charCodeAt(0) - 16384);
        }
        if (alwaysReturnAdjustedColRow) {
          return new ExceLintVector(col, row, 0);
        } else {
          return new ExceLintVector(col - origin_col, row - origin_row, 0);
        }
      }
    }

    console.log("cell is " + cell + ", origin_col = " + origin_col + ", origin_row = " + origin_row);
    throw new Error("We should never get here.");
    return ExceLintVector.Zero();
  }

  public static toR1C1(srcCell: string, destCell: string, greek = false): string {
    // Dependencies are column, then row.
    const vec1 = ExcelUtils.cell_dependency(srcCell, 0, 0);
    const vec2 = ExcelUtils.cell_dependency(destCell, 0, 0);
    let R = "R";
    let C = "C";
    if (greek) {
      // We use this encoding to avoid confusion with, say, "C1", downstream.
      R = "ρ";
      C = "γ";
    }
    // Compute the difference.
    const resultVec = vec2.subtract(vec1);
    // Now generate the R1C1 notation version, which varies
    // depending whether it's a relative or absolute reference.
    let resultStr = "";
    if (ExcelUtils.cell_both_absolute.exec(destCell)) {
      resultStr = R + vec2.y + C + vec2.x;
    } else if (ExcelUtils.cell_col_absolute.exec(destCell)) {
      if (resultVec.y === 0) {
        resultStr += R;
      } else {
        resultStr += R + "[" + resultVec[1] + "]";
      }
      resultStr += C + vec2.x;
    } else if (ExcelUtils.cell_row_absolute.exec(destCell)) {
      if (resultVec.x === 0) {
        resultStr += C;
      } else {
        resultStr += C + "[" + resultVec.x + "]";
      }
      resultStr = R + vec2.y + resultStr;
    } else {
      // Common case, both relative.
      if (resultVec.y === 0) {
        resultStr += R;
      } else {
        resultStr += R + "[" + resultVec[1] + "]";
      }
      if (resultVec.x === 0) {
        resultStr += C;
      } else {
        resultStr += C + "[" + resultVec.x + "]";
      }
    }
    return resultStr;
  }

  public static formulaToR1C1(formula: string, origin_col: number, origin_row: number): string {
    let range = formula.slice();
    const origin = ExcelUtils.column_index_to_name(origin_col) + origin_row;
    // First, get all the range pairs out.
    let found_pair: RegExpExecArray;
    while ((found_pair = ExcelUtils.range_pair.exec(range))) {
      if (found_pair) {
        range = range.replace(
          found_pair[0],
          ExcelUtils.toR1C1(origin, found_pair[1], true) + ":" + ExcelUtils.toR1C1(origin, found_pair[2], true)
        );
      }
    }

    // Now look for singletons.
    let singleton: RegExpExecArray;
    while ((singleton = ExcelUtils.single_dep.exec(range))) {
      if (singleton) {
        const first_cell = singleton[1];
        range = range.replace(singleton[0], ExcelUtils.toR1C1(origin, first_cell, true));
      }
    }
    // Now, we de-greek.
    range = range.replace(/ρ/g, "R");
    range = range.replace(/γ/g, "C");

    return range;
  }

  public static extract_sheet_cell(str: string): Array<string> {
    //	console.log("extract_sheet_cell " + str);
    const matched = ExcelUtils.sheet_plus_cell.exec(str);
    if (matched) {
      //	    console.log("extract_sheet_cell matched " + str);
      // There is only one thing to match for this pattern: we convert it into a range.
      return [matched[1], matched[2], matched[2]];
    }
    //	console.log("extract_sheet_cell failed for "+str);
    return ["", "", ""];
  }

  public static extract_sheet_range(str: string): Array<string> {
    const matched = ExcelUtils.sheet_plus_range.exec(str);
    if (matched) {
      //	    console.log("extract_sheet_range matched " + str);
      return [matched[1], matched[2], matched[3]];
    }
    //	console.log("extract_sheet_range failed to match " + str);
    return ExcelUtils.extract_sheet_cell(str);
  }

  public static make_range_string(theRange: Array<ExceLintVector>): string {
    const r = theRange;
    const col0 = r[0].x;
    const row0 = r[0].y;
    const col1 = r[1].x;
    const row1 = r[1].y;

    if (!r[0].isReference()) {
      // Not a real dependency. Skip.
      console.log("NOT A REAL DEPENDENCY: " + col1 + "," + row1);
      return "";
    } else if (col0 < 0 || row0 < 0 || col1 < 0 || row1 < 0) {
      // Defensive programming.
      console.log("WARNING: FOUND NEGATIVE VALUES.");
      return "";
    } else {
      const colname0 = ExcelUtils.column_index_to_name(col0);
      const colname1 = ExcelUtils.column_index_to_name(col1);
      //		    console.log("process: about to get range " + colname0 + row0 + ":" + colname1 + row1);
      const rangeStr = colname0 + row0 + ":" + colname1 + row1;
      return rangeStr;
    }
  }

  public static all_cell_dependencies(
    range: string,
    origin_col: number,
    origin_row: number,
    include_numbers = true
  ): ExceLintVector[] {
    let found_pair: RegExpExecArray;
    const all_vectors: ExceLintVector[] = [];

    if (typeof range !== "string") {
      return null;
    }

    // Zap all the formulas with the below characteristics.
    range = range.replace(this.formulas_with_numbers, "_"); // Don't track these.
    range = range.replace(this.formulas_with_quoted_sheetnames_2, "_");
    range = range.replace(this.formulas_with_quoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_2, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_structured_references, "_");

    /// FIX ME - should we count the same range multiple times? Or just once?

    // First, get all the range pairs out.
    while ((found_pair = ExcelUtils.range_pair.exec(range))) {
      if (found_pair) {
        const first_cell = found_pair[1];
        const first_vec = ExcelUtils.cell_dependency(first_cell, origin_col, origin_row);
        const last_cell = found_pair[2];
        const last_vec = ExcelUtils.cell_dependency(last_cell, origin_col, origin_row);

        // First_vec is the upper-left hand side of a rectangle.
        // Last_vec is the lower-right hand side of a rectangle.

        // Generate all vectors.
        const length = last_vec.x - first_vec.x + 1;
        const width = last_vec.y - first_vec.y + 1;
        for (let x = 0; x < length; x++) {
          for (let y = 0; y < width; y++) {
            all_vectors.push(new ExceLintVector(x + first_vec.x, y + first_vec.y, 0));
          }
        }

        // Wipe out the matched contents of range.
        range = range.replace(found_pair[0], "_");
      }
    }

    // Now look for singletons.
    let singleton = null;
    while ((singleton = ExcelUtils.single_dep.exec(range))) {
      if (singleton) {
        const first_cell = singleton[1];
        const vec = ExcelUtils.cell_dependency(first_cell, origin_col, origin_row);
        all_vectors.push(vec);
        // Wipe out the matched contents of range.
        range = range.replace(singleton[0], "_");
      }
    }

    if (include_numbers) {
      // Optionally roll numbers in formulas into the dependency vectors. Each number counts as "1".
      let number: RegExpExecArray;
      while ((number = ExcelUtils.number_dep.exec(range))) {
        if (number) {
          all_vectors.push(new ExceLintVector(0, 0, 1)); // just add 1 for every number
          // Wipe out the matched contents of range.
          range = range.replace(number[0], "_");
        }
      }
    }
    //	console.log("all_vectors " + originalRange + " = " + JSON.stringify(all_vectors));
    return all_vectors;
  }

  public static numeric_constants(range: string): Array<number> {
    const numbers = [];
    range = range.slice();
    if (typeof range !== "string") {
      return numbers;
    }

    // Zap all the formulas with the below characteristics.
    range = range.replace(this.formulas_with_numbers, "_"); // Don't track these.
    range = range.replace(this.formulas_with_quoted_sheetnames_2, "_");
    range = range.replace(this.formulas_with_quoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_2, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_unquoted_sheetnames_1, "_");
    range = range.replace(this.formulas_with_structured_references, "_");

    // First, get all the range pairs out.
    let found_pair: RegExpExecArray;
    while ((found_pair = ExcelUtils.range_pair.exec(range))) {
      if (found_pair) {
        // Wipe out the matched contents of range.
        range = range.replace(found_pair[0], "_");
      }
    }

    // Now look for singletons.
    let singleton: RegExpExecArray;
    while ((singleton = ExcelUtils.single_dep.exec(range))) {
      if (singleton) {
        // Wipe out the matched contents of range.
        range = range.replace(singleton[0], "_");
      }
    }

    // Now aggregate total numeric constants (sum them).
    let number: RegExpExecArray;
    //	let total = 0.0;
    while ((number = ExcelUtils.number_dep.exec(range))) {
      if (number) {
        numbers.push(parseFloat(number[0]));
        //		total += parseFloat(number);
        // Wipe out the matched contents of range.
        range = range.replace(number[0], "_");
      }
    }
    return numbers; // total;
  }

  public static baseVector(): ExceLintVector {
    return new ExceLintVector(0, 0, 0);
  }

  public static all_dependencies(
    row: number,
    col: number,
    origin_row: number,
    origin_col: number,
    formulas: Spreadsheet
  ): ExceLintVector[] {
    // Discard references to cells outside the formula range.
    if (row >= formulas.length || col >= formulas[0].length || row < 0 || col < 0) {
      return [];
    }

    // Check if this cell is a formula.
    const cell = formulas[row][col];
    if (cell.length > 1 && cell[0] === "=") {
      // It is. Compute the dependencies.
      return ExcelUtils.all_cell_dependencies(cell, origin_col, origin_row);
    } else {
      return [];
    }
  }

  // This function returns a dictionary (Dict<boolean>)) of all of the addresses
  // that are referenced by some formula, where the key is the address and the
  // value is always the boolean true.
  public static generate_all_references(formulas: Spreadsheet, origin_col: number, origin_row: number): Dict<boolean> {
    // initialize dictionary
    const refs: Dict<boolean> = {};

    let counter = 0;
    for (let i = 0; i < formulas.length; i++) {
      const row = formulas[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        counter++;
        if (counter % 1000 === 0) {
          //		    console.log(counter + " references down");
        }

        if (cell[0] === "=") {
          // It's a formula.
          const direct_refs = ExcelUtils.all_cell_dependencies(cell, 0, 0); // origin_col, origin_row); // was just 0,0....  origin_col, origin_row);
          for (const dep of direct_refs) {
            if (!dep.isReference()) {
              // Not a real reference. Skip.
            } else {
              // Check to see if this is data or a formula.
              // If it's not a formula, add it.
              const rowIndex = dep.x - origin_col - 1;
              const colIndex = dep.y - origin_row - 1;
              const outsideFormulaRange =
                colIndex >= formulas.length || rowIndex >= formulas[0].length || rowIndex < 0 || colIndex < 0;
              let addReference = false;
              if (outsideFormulaRange) {
                addReference = true;
              } else {
                // Only include non-formulas (if they are in the range).
                const referentCell = formulas[colIndex][rowIndex];
                if (referentCell !== undefined && referentCell[0] !== "=") {
                  addReference = true;
                }
              }
              if (addReference) {
                refs[dep.asKey()] = true;
              }
            }
          }
        }
      }
    }
    return refs;
  }

  /**
   * Converts an A1 address string into an R1C1 tuple where the first
   * element is the column number and the second element is the row number.
   * @param a1addr An address string in A1 format
   */
  public static addrA1toR1C1(a1addr: string): Address {
    // split sheet name, remove absolute reference symbols, and
    // ensure address is uppercase
    const a1normed = a1addr.replace("$", "");
    const aa = a1normed.split("!");
    const sheet = aa[0];
    const addr = aa[1].toUpperCase();
    let processCol = true;

    // accumulated characters go here
    const x_list: number[] = [];
    const y_list: number[] = [];

    for (let i = 0; i < addr.length; i++) {
      const c = addr.charAt(i);

      // process the column
      if (processCol) {
        const n = Number(c);
        if (!isNaN(n)) {
          // switch to processing y once we see numeric chars
          processCol = false;
        } else {
          // e.g., A = ASCII decimal 65, so A will equal 1, Z will equal 26.
          const code = c.charCodeAt(0);
          x_list.push(code - 64);
        }
      }

      // process the row
      if (!processCol) {
        // the magnitude of _y depends on how many y chars there are,
        // which we don't yet know.  Keep a count and do the math later.
        y_list.push(Number(c));
      }
    }

    // so that we can process from the least significant digit
    x_list.reverse();
    y_list.reverse();
    const col = x_list.map((t, i) => t * Math.pow(26, i)).reduce((acc, e) => acc + e, 0);
    const row = y_list.map((t, i) => t * Math.pow(10, i)).reduce((acc, e) => acc + e, 0);
    return new Address(sheet, row, col);
  }
}
