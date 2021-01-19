const fs = require("fs");
const textdiff = require("text-diff");
const diff = new textdiff();

import { ExcelUtils } from "./excelutils";

class CellEncoder {
  private static maxRows = 64; // -32..32
  private static maxColumns = 32; // -16..16
  private static absoluteRowMultiplier: number =
    2 * CellEncoder.maxRows * CellEncoder.maxColumns; // if bit set, absolute row
  public static absoluteColumnMultiplier: number =
    2 * CellEncoder.absoluteRowMultiplier; // if bit set, absolute column

  public static startPoint = 2048; // Start the encoding of the cell at this Unicode value

  public static encode(
    col: number,
    row: number,
    absoluteColumn = false,
    absoluteRow = false
  ): number {
    const addAbsolutes =
      Number(absoluteRow) * CellEncoder.absoluteRowMultiplier +
      Number(absoluteColumn) * CellEncoder.absoluteColumnMultiplier;
    return (
      addAbsolutes +
      CellEncoder.maxRows * (CellEncoder.maxColumns / 2 + col) +
      (CellEncoder.maxRows / 2 + row) +
      CellEncoder.startPoint
    );
  }

  public static encodeToChar(
    col: number,
    row: number,
    absoluteColumn = false,
    absoluteRow = false
  ): string {
    const chr = String.fromCodePoint(
      CellEncoder.encode(col, row, absoluteColumn, absoluteRow)
    );
    return chr;
  }

  private static decodeColumn(encoded: number): number {
    encoded -= CellEncoder.startPoint;
    return (
      Math.floor(encoded / CellEncoder.maxRows) - CellEncoder.maxColumns / 2
    );
  }

  private static decodeRow(encoded: number): number {
    encoded -= CellEncoder.startPoint;
    return (encoded % CellEncoder.maxRows) - CellEncoder.maxRows / 2;
  }

  public static decodeFromChar(
    chr: string
  ): [number, number, boolean, boolean] {
    let decodedNum = chr.codePointAt(0);
    let absoluteColumn = false;
    let absoluteRow = false;
    if (decodedNum & CellEncoder.absoluteRowMultiplier) {
      decodedNum &= ~CellEncoder.absoluteRowMultiplier;
      absoluteRow = true;
    }
    if (decodedNum & CellEncoder.absoluteColumnMultiplier) {
      decodedNum &= ~CellEncoder.absoluteColumnMultiplier;
      absoluteColumn = true;
    }
    const result: [number, number, boolean, boolean] = [
      CellEncoder.decodeColumn(decodedNum),
      CellEncoder.decodeRow(decodedNum),
      absoluteColumn,
      absoluteRow,
    ];
    return result;
  }

  public static maxEncodedSize(): number {
    return (
      CellEncoder.encode(CellEncoder.maxColumns - 1, CellEncoder.maxRows - 1) -
      CellEncoder.encode(
        -(CellEncoder.maxColumns - 1),
        -(CellEncoder.maxRows - 1)
      )
    );
  }

  public static test(): void {
    for (
      let col = -CellEncoder.maxColumns;
      col < CellEncoder.maxColumns;
      col++
    ) {
      for (let row = -CellEncoder.maxRows; row < CellEncoder.maxRows; row++) {
        const encoded = CellEncoder.encode(col, row);
        const decodedCol = CellEncoder.decodeColumn(encoded);
        const decodedRow = CellEncoder.decodeRow(encoded);
        //	console.log(decodedCol + " " + decodedRow);
        console.assert(col === decodedCol, "NOPE COL");
        console.assert(row === decodedRow, "NOPE ROW");
      }
    }
  }
}

export class FixDiff {
  // Load the JSON file containing all the Excel functions.
  private fns = JSON.parse(fs.readFileSync("functions.json", "utf-8"));
  // console.log(JSON.stringify(fns));

  // Build a map of Excel functions to crazy Unicode characters and back
  // again.  We do this so that diffs are in effect "token by token"
  // (e.g., so ROUND and RAND don't get a diff of "O|A" but are instead
  // considered entirely different).
  private fn2unicode = {};
  private unicode2fn = {};

  // Construct the arrays above.
  private initArrays() {
    let i = 0;
    for (const fnName of this.fns) {
      const str = String.fromCharCode(256 + i);
      this.fn2unicode[fnName] = str;
      this.unicode2fn[str] = fnName;
      i++;
      // console.log(fnName + " " + this.fn2unicode[fnName] + " " + this.unicode2fn[this.fn2unicode[fnName]]);
    }

    // Sort the functions in reverse order by size (longest first). This
    // order will prevent accidentally tokenizing substrings of functions.
    this.fns.sort((a, b) => {
      if (a.length < b.length) {
        return 1;
      }
      if (a.length > b.length) {
        return -1;
      } else {
        // Sort in alphabetical order.
        if (a < b) {
          return -1;
        }
        if (a > b) {
          return 1;
        }
        return 0;
      }
    });
  }

  constructor() {
    this.initArrays();
  }

  public static toPseudoR1C1(srcCell: string, destCell: string): string {
    // Dependencies are column, then row.
    const vec1 = ExcelUtils.cell_dependency(srcCell, 0, 0);
    const vec2 = ExcelUtils.cell_dependency(destCell, 0, 0);
    console.log("start " + JSON.stringify(vec1));
    console.log("dest  " + JSON.stringify(vec2));
    // Compute the difference.
    const resultVec = vec2.subtract(vec1);
    console.log("vec2  " + JSON.stringify(resultVec));
    // Now generate the pseudo R1C1 version, which varies
    // depending whether it's a relative or absolute reference.
    let resultStr = "";
    if (ExcelUtils.cell_both_absolute.exec(destCell)) {
      console.log("both absolute");
      resultStr = CellEncoder.encodeToChar(vec2[0], vec2[1], true, true);
    } else if (ExcelUtils.cell_col_absolute.exec(destCell)) {
      console.log("column absolute, row relative");
      console.log(vec2[0]);
      console.log(resultVec[1]);
      resultStr = CellEncoder.encodeToChar(vec2[0], resultVec[1], true, false);
    } else if (ExcelUtils.cell_row_absolute.exec(destCell)) {
      console.log("row absolute, column relative");
      resultStr = CellEncoder.encodeToChar(resultVec[0], vec2[1], false, true);
    } else {
      // Common case, both relative.
      console.log("both relative");
      resultStr = CellEncoder.encodeToChar(
        resultVec[0],
        resultVec[1],
        false,
        false
      );
    }
    console.log("to pseudo r1c1: " + resultStr);
    return resultStr;
  }

  public static formulaToPseudoR1C1(
    formula: string,
    origin_col: number,
    origin_row: number
  ): string {
    let range = formula.slice();
    const origin = ExcelUtils.column_index_to_name(origin_col) + origin_row;
    // First, get all the range pairs out.
    let found_pair;
    while ((found_pair = ExcelUtils.range_pair.exec(range))) {
      if (found_pair) {
        const first_cell = found_pair[1];
        const last_cell = found_pair[2];
        range = range.replace(
          found_pair[0],
          FixDiff.toPseudoR1C1(origin, found_pair[1]) +
            ":" +
            FixDiff.toPseudoR1C1(origin, found_pair[2])
        );
      }
    }

    // Now look for singletons.
    let singleton = null;
    while ((singleton = ExcelUtils.single_dep.exec(range))) {
      if (singleton) {
        const first_cell = singleton[1];
        range = range.replace(
          singleton[0],
          FixDiff.toPseudoR1C1(origin, first_cell)
        );
      }
    }
    return range;
  }

  public tokenize(formula: string): string {
    for (let i = 0; i < this.fns.length; i++) {
      formula = formula.replace(this.fns[i], this.fn2unicode[this.fns[i]]);
    }
    formula = formula.replace(/(\-?\d+)/g, (_, num) => {
      // Make sure the unicode characters are far away from the encoded cell values.
      const replacement = String.fromCodePoint(
        CellEncoder.absoluteColumnMultiplier * 2 + parseInt(num)
      );
      return replacement;
    });
    return formula;
  }

  public detokenize(formula: string): string {
    for (let i = 0; i < this.fns.length; i++) {
      formula = formula.replace(this.fn2unicode[this.fns[i]], this.fns[i]);
    }
    return formula;
  }

  // Return the diffs (with formulas treated specially).
  public compute_fix_diff(str1, str2, c1, r1, c2, r2) {
    //c2, r2) {
    // Convert to pseudo R1C1 format.
    let rc_str1 = FixDiff.formulaToPseudoR1C1(str1, c1, r1); // ExcelUtils.formulaToR1C1(str1, c1, r1);
    let rc_str2 = FixDiff.formulaToPseudoR1C1(str2, c2, r2); // ExcelUtils.formulaToR1C1(str2, c2, r2);
    // Tokenize the functions.
    rc_str1 = this.tokenize(rc_str1);
    rc_str2 = this.tokenize(rc_str2);
    // Build up the diff.
    const difference = diff.main(rc_str1, rc_str2);
    const theDiff = [[...difference], [...difference]];
    // Now de-tokenize the diff contents
    // and convert back out of pseudo R1C1 format.
    for (let i = 0; i < 2; i++) {
      for (let j = 0; j < theDiff[i].length; j++) {
        if (theDiff[i][j][0] === 0) {
          // No diff
          if (i === 0) {
            theDiff[i][j][1] = this.fromPseudoR1C1(theDiff[i][j][1], c1, r1); // first one
          } else {
            theDiff[i][j][1] = this.fromPseudoR1C1(theDiff[i][j][1], c2, r2); // second one
          }
        } else if (theDiff[i][j][0] === -1) {
          // Left diff
          theDiff[i][j][1] = this.fromPseudoR1C1(theDiff[i][j][1], c1, r1);
        } else {
          // Right diff
          theDiff[i][j][1] = this.fromPseudoR1C1(theDiff[i][j][1], c2, r2);
        }
        theDiff[i][j][1] = this.detokenize(theDiff[i][j][1]);
      }
      diff.cleanupSemantic(theDiff[i]);
    }
    return theDiff;
  }

  public fromPseudoR1C1(
    r1c1_formula: string,
    origin_col: number,
    origin_row: number
  ): string {
    // We assume that formulas have already been 'tokenized'.
    // console.log("fromPseudoR1C1 = " + r1c1_formula + ", origin_col = " + origin_col + ", origin_row = " + origin_row);
    let r1c1 = r1c1_formula.slice();
    // Find the Unicode characters and decode them.
    r1c1 = r1c1.replace(/([\u800-\uF000])/g, (_full, encoded_char) => {
      if (encoded_char.codePointAt(0) < CellEncoder.startPoint) {
        return encoded_char;
      }
      const [co, ro, absCo, absRo] = CellEncoder.decodeFromChar(encoded_char);
      let result: string;
      if (!absCo && !absRo) {
        // Both relative (R[..]C[...])
        // console.log("both relative");
        result =
          ExcelUtils.column_index_to_name(origin_col + co) + (origin_row + ro);
      }
      if (absCo && !absRo) {
        // Row relative, column absolute (R[..]C...)
        // console.log("column absolute");
        result = "$" + ExcelUtils.column_index_to_name(co) + (origin_row + ro);
      }
      if (!absCo && absRo) {
        // Row absolute, column relative (R...C[..])
        // console.log("row absolute");
        result = ExcelUtils.column_index_to_name(origin_col + co) + "$" + ro;
      }
      if (absCo && absRo) {
        // Both absolute (R...C...)
        // console.log("both absolute");
        result = "$" + ExcelUtils.column_index_to_name(co) + "$" + ro;
      }
      return result;
    });
    return r1c1;
  }

  private static redtext = "\u001b[31m";
  private static yellowtext = "\u001b[33m";
  private static greentext = "\u001b[32m";
  private static whitetext = "\u001b[37m";
  private static resettext = "\u001b[0m";
  private static textcolor = [
    FixDiff.redtext,
    FixDiff.yellowtext,
    FixDiff.greentext,
  ];

  public pretty_diffs(diffs): string[] {
    const strList = [];
    // Iterate for -1 and 1.
    for (const i of [-1, 1]) {
      // console.log(i);
      let str = "";
      for (const d of diffs) {
        // console.log("diff = " + JSON.stringify(d));
        if (d[0] === i) {
          str += FixDiff.textcolor[i + 1] + d[1] + FixDiff.resettext;
        } else if (d[0] === 0) {
          str += FixDiff.whitetext + d[1] + FixDiff.resettext;
        }
      }
      strList.push(str);
    }
    return strList;
  }
}

function showDiffs(str1, row1, col1, str2, row2, col2) {
  const nd = new FixDiff();
  const [diff0, diff1] = nd.compute_fix_diff(
    str1,
    str2,
    col1 - 1,
    row1 - 1,
    col2 - 1,
    row2 - 1
  );
  console.log(diff0);
  console.log(diff1);
  // console.log(JSON.stringify(diffs));
  const [first, dummy2] = nd.pretty_diffs(diff0);
  const [dummy1, second] = nd.pretty_diffs(diff1);
  // the first one should be the one needing to be fixed.

  //    console.log(str1);
  //    console.log(str2);
  //    console.log("change: ");
  console.log(first);
  console.log(dummy2);

  console.log(dummy1);
  console.log(second);
}

showDiffs("=ROUND(C9:E$10)", 1, 3, "=ROUND(B9:E10)", 1, 2);
showDiffs("=ROUND(B9:E10)", 1, 2, "=ROUND(C9:E$10)", 1, 3);
showDiffs("=ROUND(B9:E10)", 1, 2, "=ROUND(C9:E10)", 1, 3);
showDiffs("=ROUND($B$9:E10)", 1, 2, "=ROUND(C9:E10)", 1, 3);

// Now try a diff.
const [row1, col1] = [1, 3];
const [row2, col2] = [1, 2];
//let [row1, col1] = [11, 2];
//let [row2, col2] = [11, 3];
//let str1 = '=ROUND(B7:B9)'; // 'ROUND(A1)+12';
//let str2 = '=ROUND(C7:C10)'; // 'ROUNDUP(B2)+12';
const str1 = "=ROUND(E$10:C9)"; // 'ROUNDUP(B2)+12';
const str2 = "=ROUND(E10:B9)"; // 'ROUND(A1)+12';
