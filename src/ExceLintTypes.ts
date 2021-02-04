import { WorksheetOutput, WorkbookOutput } from "./exceljson";
import { ExcelUtils } from "./excelutils";
import { Classification } from "./classification";
import { Config } from "./config";
import { Option, Some, None, flatMap } from "./option";

export interface Dict<V> {
  [key: string]: V;
}

// all users of Spreadsheet store their data in row-major format (i.e., indexed by y first, then x).
export type Spreadsheet = string[][];

export type Fingerprint = string;

export type Metric = number;

// a rectangle is defined by its start and end vectors
export type Rectangle = [ExceLintVector, ExceLintVector];

export class ProposedFix {
  // This comment no longer holds, since there is a data type for ProposedFix,
  // but it is informative, since it explains the meaning of the fields:
  // ## old comment ##
  //   Format of proposed fixes =, e.g., [-3.016844756293869, [[5,7],[5,11]],[[6,7],[6,11]]]
  //   entropy, and two ranges:
  //      upper-left corner of range (column, row), lower-right corner of range (column, row)
  // ## end old comment ##
  private _score: number; // fix distance (entropy)
  private _rect1: Rectangle; // suspected bug
  private _rect2: Rectangle; // merge candidate
  private _sameFormat: boolean = true; // the two rectangles have the same format
  private _analysis: Option<FixAnalysis> = None; // we add this later, after we analyze the fix

  constructor(score: number, rect1: Rectangle, rect2: Rectangle) {
    this._score = score;
    this._rect1 = rect1;
    this._rect2 = rect2;
  }

  public get rectangles(): [Rectangle, Rectangle] {
    return [this._rect1, this.rect2];
  }

  public get score(): number {
    return this._score;
  }

  public set score(s: number) {
    this._score = s;
  }

  public get rect1(): Rectangle {
    return this._rect1;
  }

  public get rect2(): Rectangle {
    return this._rect2;
  }

  public get analysis(): FixAnalysis {
    if (this._analysis.hasValue) {
      return this._analysis.value;
    } else {
      throw new Error("Cannot obtain analysis about unanalyzed fix.");
    }
  }

  public set analysis(fix_analysis: FixAnalysis) {
    this._analysis = new Some(fix_analysis);
  }

  public get sameFormat(): boolean {
    return this._sameFormat;
  }

  public set sameFormat(is_same: boolean) {
    this._sameFormat = is_same;
  }

  public equals(other: ProposedFix): boolean {
    const [this_r1_ul, this_r1_br] = this.rect1;
    const [this_r2_ul, this_r2_br] = this.rect2;

    const [other_r1_ul, other_r1_br] = other.rect1;
    const [other_r2_ul, other_r2_br] = other.rect2;

    return (
      this_r1_ul.equals(other_r1_ul) &&
      this_r1_br.equals(other_r1_br) &&
      this_r2_ul.equals(other_r2_ul) &&
      this_r2_br.equals(other_r2_br) &&
      this.score === other.score
    );
  }
}

export function upperleft(r: Rectangle): ExceLintVector {
  return r[0];
}

export function bottomright(r: Rectangle): ExceLintVector {
  return r[1];
}

// Convert a rectangle into a list of vectors.
export function expand(first: ExceLintVector, second: ExceLintVector): ExceLintVector[] {
  const expanded: ExceLintVector[] = [];
  for (let i = first.x; i <= second.x; i++) {
    for (let j = first.y; j <= second.y; j++) {
      expanded.push(new ExceLintVector(i, j, 0));
    }
  }
  return expanded;
}

export class ExceLintVector {
  public x: number;
  public y: number;
  public c: number;

  constructor(x: number, y: number, c: number) {
    this.x = x;
    this.y = y;
    this.c = c;
  }

  public isConstant(): boolean {
    return this.c === 1;
  }

  public static Zero(): ExceLintVector {
    return new ExceLintVector(0, 0, 0);
  }

  // Subtract other from this vector
  public subtract(v: ExceLintVector): ExceLintVector {
    return new ExceLintVector(this.x - v.x, this.y - v.y, this.c - v.c);
  }

  // Turn this vector into a string that can be used as a dictionary key
  public asKey(): string {
    return this.x.toString() + "," + this.y.toString() + "," + this.c.toString();
  }

  // Return true if this vector encodes a reference
  public isReference(): boolean {
    return !(this.x === 0 && this.y === 0 && this.c !== 0);
  }

  // Pretty-print vectors
  public toString(): string {
    return "<" + this.asKey() + ">";
  }

  // performs a deep eqality check
  public equals(other: ExceLintVector): boolean {
    return this.x === other.x && this.y === other.y && this.c === other.c;
  }

  public hash(): number {
    // This computes a weighted L1 norm of the vector
    const v0 = Math.abs(this.x);
    const v1 = Math.abs(this.y);
    const v2 = this.c;
    return ExceLintVector.Multiplier * (v0 + v1 + v2);
  }

  // vector sum reduction
  public static readonly VectorSum = (acc: ExceLintVector, curr: ExceLintVector): ExceLintVector =>
    new ExceLintVector(acc.x + curr.x, acc.y + curr.y, acc.c + curr.c);

  public static vectorSetEquals(set1: ExceLintVector[], set2: ExceLintVector[]): boolean {
    // create a hashs et with elements from set1,
    // and then check that set2 induces the same set
    const hset: Set<number> = new Set();
    set1.forEach((v) => hset.add(v.hash()));

    // check hset1 for hashes of elements in set2.
    // if there is a match, remove the element from hset1.
    // if there isn't a match, return early.
    for (let i = 0; i < set2.length; i++) {
      const h = set2[i].hash();
      if (hset.has(h)) {
        hset.delete(h);
      } else {
        // sets are definitely not equal
        return false;
      }
    }

    // sets are equal iff hset has no remaining elements
    return hset.size === 0;
  }

  // A multiplier for the hash function.
  public static readonly Multiplier = 1; // 103037;

  // Given an array of ExceLintVectors, returns an array of unique
  // ExceLintVectors.  This explicitly does not return Javascript's Set
  // datatype, which is inherently dangerous for UDTs, since it curiously
  // provides no mechanism for specifying membership based on user-defined
  // object equality.
  public static toSet(vs: ExceLintVector[]): ExceLintVector[] {
    const out: ExceLintVector[] = [];
    const hset: Set<number> = new Set();
    for (const i in vs) {
      const v = vs[i];
      const h = v.hash();
      if (!hset.has(h)) {
        out.push(v);
        hset.add(h);
      }
    }
    return out;
  }
}

export class Analysis {
  suspicious_cells: ExceLintVector[];
  grouped_formulas: Dict<Rectangle[]>;
  grouped_data: Dict<Rectangle[]>;
  proposed_fixes: ProposedFix[];

  constructor(
    suspicious_cells: ExceLintVector[],
    grouped_formulas: Dict<Rectangle[]>,
    grouped_data: Dict<Rectangle[]>,
    proposed_fixes: ProposedFix[]
  ) {
    this.suspicious_cells = suspicious_cells;
    this.grouped_formulas = grouped_formulas;
    this.grouped_data = grouped_data;
    this.proposed_fixes = proposed_fixes;
  }
}

export function vectorComparator(v1: ExceLintVector, v2: ExceLintVector): number {
  if (v1.x < v2.x) {
    return -1;
  }
  if (v1.x > v2.x) {
    return 1;
  }
  if (v1.y < v2.y) {
    return -1;
  }
  if (v1.y > v2.y) {
    return 1;
  }
  if (v1.c < v2.c) {
    return -1;
  }
  if (v1.c > v2.c) {
    return 1;
  }
  return 0;
}

// A comparator that sorts rectangles by their upper-left and then lower-right
// vectors.
export function rectangleComparator(r1: Rectangle, r2: Rectangle): number {
  const cmp = vectorComparator(r1[0], r2[0]);
  if (cmp == 0) {
    return vectorComparator(r1[1], r2[1]);
  } else {
    return cmp;
  }
}

export class RectInfo {
  formula: string; // actual formula
  constants: number[] = []; // all the numeric constants in each formula
  sum: number; // the sum of all the numeric constants in each formula
  dependencies: ExceLintVector[] = []; // the set of no-constant dependence vectors in the formula
  dependence_count: number; // the number of dependent cells
  absolute_refcount: number; // the number of absolute references in each formula
  r1c1_formula: string; // formula in R1C1 format
  r1c1_print_formula: string; // as above, but for R1C1 formulas
  print_formula: string; // formula with a preface (the cell name containing each)

  constructor(rect: Rectangle, sheet: WorksheetOutput) {
    // the coordinates of the cell containing the first formula in the proposed fix range
    const formulaCoord = rect[0];
    const y = formulaCoord.y - 1; // row
    const x = formulaCoord.x - 1; // col
    this.formula = sheet.formulas[y][x]; // the formula itself
    this.constants = ExcelUtils.numeric_constants(this.formula); // all numeric constants in the formula
    this.sum = this.constants.reduce((a, b) => a + b, 0); // the sum of all numeric constants
    this.dependencies = ExcelUtils.all_cell_dependencies(this.formula, x + 1, y + 1, false);
    this.dependence_count = this.dependencies.length;
    this.absolute_refcount = (this.formula.match(/\$/g) || []).length;
    this.r1c1_formula = ExcelUtils.formulaToR1C1(this.formula, x + 1, y + 1);
    const preface = ExcelUtils.column_index_to_name(x + 1) + (y + 1) + ":";
    this.r1c1_print_formula = preface + this.r1c1_formula;
    this.print_formula = preface + this.formula;
  }
}

export class FixAnalysis {
  classification: Classification.BinCategory[];
  analysis: RectInfo[];
  direction_is_vert: boolean;

  constructor(
    classification: Classification.BinCategory[],
    analysis: RectInfo[],
    direction_is_vert: boolean
  ) {
    this.classification = classification;
    this.analysis = analysis;
    this.direction_is_vert = direction_is_vert;
  }

  // Compute the difference in constant sums
  public totalNumericDifference(): number {
    return Math.abs(this.analysis[0].sum - this.analysis[1].sum);
  }

  // Compute the magnitude of the difference in constant sums
  public magnitudeNumericDifference(): number {
    const n = this.totalNumericDifference();
    return n === 0 ? 0 : Math.log10(n);
  }
}

export class WorkbookAnalysis {
  private sheets: Dict<WorksheetAnalysis> = {};

  public getSheet(name: string) {
    return this.sheets[name];
  }

  public addSheet(s: WorksheetAnalysis) {
    this.sheets[s.name] = s;
  }
}

export class WorksheetAnalysis {
  private readonly sheet: WorksheetOutput;
  private readonly pf: ProposedFix[];
  private readonly foundBugs: ExceLintVector[];
  private readonly analysis: Analysis;

  constructor(sheet: WorksheetOutput, pf: ProposedFix[], a: Analysis) {
    this.sheet = sheet;
    this.pf = pf;
    this.foundBugs = WorksheetAnalysis.createBugList(pf);
    this.analysis = a;
  }

  // Get the grouped data
  get groupedData(): Dict<Rectangle[]> {
    return this.analysis.grouped_data;
  }

  // Get the grouped formulas
  get groupedFormulas(): Dict<Rectangle[]> {
    return this.analysis.grouped_formulas;
  }

  // Get the sheet name
  get name(): string {
    return this.sheet.sheetName;
  }

  // Get all of the proposed fixes.
  get proposedFixes(): ProposedFix[] {
    return this.pf;
  }

  // Compute number of cells containing formulas.
  get numFormulaCells(): number {
    return this.sheet.formulas.flat().filter((x) => x.length > 0).length;
  }

  // Count the number of non-empty cells.
  get numValueCells(): number {
    return this.sheet.values.flat().filter((x) => x.length > 0).length;
  }

  // Compute number of columns
  get columns(): number {
    return this.sheet.values[0].length;
  }

  // Compute number of rows
  get rows(): number {
    return this.sheet.values.length;
  }

  // Compute total number of cells
  get totalCells(): number {
    return this.rows * this.columns;
  }

  // Produce a sum total of all of the entropy scores for use as a weight
  get weightedAnomalousRanges(): number {
    return this.pf.map((x) => x[0]).reduce((x, y) => x + y, 0);
  }

  // Get the total number of anomalous cells
  get numAnomalousCells(): number {
    return this.foundBugs.length;
  }

  // For every proposed fix, if it is above the score threshold, keep it,
  // and return the unique set of all vectors contained in any kept fix.
  private static createBugList(pf: ProposedFix[]): ExceLintVector[] {
    const keep: ExceLintVector[][] = flatMap((pf) => {
      if (pf.score >= Config.reportingThreshold / 100) {
        const rect1cells = expand(upperleft(pf.rect1), bottomright(pf.rect1));
        const rect2cells = expand(upperleft(pf.rect2), bottomright(pf.rect2));
        return new Some(rect1cells.concat(rect2cells));
      } else {
        return None;
      }
    }, pf);
    let flattened = keep.flat(1);
    return ExceLintVector.toSet(flattened);
  }
}
