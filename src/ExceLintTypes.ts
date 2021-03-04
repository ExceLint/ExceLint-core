import { WorkbookOutput, WorksheetOutput } from "./exceljson";
import { ExcelUtils } from "./excelutils";
import { Classification } from "./classification";
import { Config } from "./config";
import { Option, Some, None, flatMap } from "./option";

interface Dict<V> {
  [key: string]: V;
}

export class Dictionary<V> {
  private _d: Dict<V> = {};
  public contains(key: string): boolean {
    return this._d[key] !== undefined;
  }
  public get(key: string): V {
    if (this.contains(key)) {
      return this._d[key];
    } else {
      throw new Error("Cannot get unknown key '" + key + "' in dictionary.");
    }
  }
  public put(key: string, value: V): void {
    this._d[key] = value;
  }
  public del(key: string): V {
    if (this.contains(key)) {
      const v = this._d[key];
      delete this._d[key];
      return v;
    } else {
      throw new Error("Cannot delete unknown key '" + key + "' in dictionary.");
    }
  }
  public get keys(): string[] {
    const output: string[] = [];
    for (let key in this._d) {
      output.push(key);
    }
    return output;
  }
  /**
   * Performs a shallow copy of the dictionary.
   */
  public clone(): Dictionary<V> {
    const dict = new Dictionary<V>();
    for (const key of this.keys) {
      dict.put(key, this.get(key));
    }
    return dict;
  }
}

export interface IComparable<V> {
  equals(v: IComparable<V>): boolean;
}

export class CSet<V extends IComparable<V>> implements IComparable<CSet<V>> {
  private _vs: V[] = [];

  constructor(values: V[]) {
    for (let i = 0; i < values.length; i++) {
      this.add(values[i]);
    }
  }

  public add(v: V): boolean {
    let keep = true;
    for (let i = 0; i < this._vs.length; i++) {
      if (v.equals(this._vs[i])) {
        keep = false;
        break;
      }
    }
    if (keep) this._vs.push(v);
    return keep;
  }

  public get size() {
    return this._vs.length;
  }

  public get values() {
    return this._vs;
  }

  private clone(): CSet<V> {
    return new CSet(this.values);
  }

  public equals(vs: CSet<V>) {
    if (this.size !== vs.size) {
      return false;
    }
    const copy = this.clone();
    const values = vs.values;
    for (let i = 0; i < values.length; i++) {
      copy.add(values[i]);
      if (this.size !== copy.size) return false;
    }
    return true;
  }

  public map<X extends IComparable<X>>(f: (v: V) => X): CSet<X> {
    const output = CSet.empty<X>();
    for (let i = 0; i < this._vs.length; i++) {
      const fi = f(this._vs[i]);
      output.add(fi);
    }
    return output;
  }

  /**
   * Returns a new set object which is the union
   * @param set
   */
  public union(set: CSet<V>): CSet<V> {
    const output = this.clone();
    for (let i = 0; i < set.values.length; i++) {
      output.add(set.values[i]);
    }
    return output;
  }

  public static empty<T extends IComparable<T>>(): CSet<T> {
    return new CSet<T>([]);
  }

  public toString(): string {
    return "{" + this._vs.join(",") + "}";
  }
}

// all users of Spreadsheet store their data in row-major format (i.e., indexed by y first, then x).
export type Spreadsheet = string[][];

export class Address implements IComparable<Address> {
  private _sheet: string;
  private _row: number;
  private _column: number;
  constructor(sheet: string, row: number, column: number) {
    this._sheet = sheet;
    this._row = row;
    this._column = column;
  }
  public get row(): number {
    return this._row;
  }
  public get column(): number {
    return this._column;
  }
  public get worksheet(): string {
    return this._sheet;
  }
  public equals(a: Address): boolean {
    return this._sheet === a._sheet && this._row === a._row && this._column === a._column;
  }
  public toString(): string {
    return "R" + this._column + "C" + this._row;
  }
}

export class Fingerprint implements IComparable<Fingerprint> {
  private _fp: number;

  constructor(fpval: number) {
    this._fp = fpval;
  }

  public equals(f: Fingerprint): boolean {
    return this._fp === f._fp;
  }

  public asKey(): string {
    return this._fp.toString();
  }

  public static fromKey(key: string): Fingerprint {
    return new Fingerprint(parseInt(key));
  }
}

export type Metric = number;

// a rectangle is defined by its start and end vectors
export class Rectangle implements IComparable<Rectangle> {
  private _tl: ExceLintVector;
  private _br: ExceLintVector;

  constructor(tl: ExceLintVector, br: ExceLintVector) {
    this._tl = tl;
    this._br = br;
  }

  public equals(r: Rectangle): boolean {
    return this._tl.equals(r._tl) && this._br.equals(r._br);
  }

  public get upperleft() {
    return this._tl;
  }

  public get bottomright() {
    return this._br;
  }

  public expand(): ExceLintVector[] {
    return expand(this._tl, this._br);
  }
}

export class ProposedFix implements IComparable<ProposedFix> {
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
    return this._rect1.equals(other._rect1) && this._rect2.equals(other._rect2);
  }

  public includesCellAt(addr: Address): boolean {
    // convert addr to an ExceLintVector
    const v = new ExceLintVector(addr.column, addr.row, 0);

    // check the rectangles
    const first_cells = this.rect1.expand();
    const second_cells = this.rect2.expand();
    const all_cells = first_cells.concat(second_cells);
    for (let i = 0; i < all_cells.length; i++) {
      const cell = all_cells[i];
      if (v.equals(cell)) {
        return true;
      }
    }
    return false;
  }
}

export function upperleft(r: Rectangle): ExceLintVector {
  return r.upperleft;
}

export function bottomright(r: Rectangle): ExceLintVector {
  return r.bottomright;
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

  // Turn a key into a vector
  public static fromKey(key: string): ExceLintVector {
    const parts = key.split(",");
    const x = parseInt(parts[0]);
    const y = parseInt(parts[1]);
    const c = parseInt(parts[2]);
    return new ExceLintVector(x, y, c);
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

  /**
   * Gets the vector of the cell above this one
   */
  public get up(): ExceLintVector {
    return new ExceLintVector(this.x, this.y - 1, this.c);
  }

  /**
   * Gets the vector of the cell below this one
   */
  public get down(): ExceLintVector {
    return new ExceLintVector(this.x, this.y + 1, this.c);
  }

  /**
   * Gets the vector of the cell to the left of this one
   */
  public get left(): ExceLintVector {
    return new ExceLintVector(this.x - 1, this.y, this.c);
  }

  /**
   * Gets the vector of the cell to the right of this one
   */
  public get right(): ExceLintVector {
    return new ExceLintVector(this.x + 1, this.y, this.c);
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
  grouped_formulas: Dictionary<Rectangle[]>;
  grouped_data: Dictionary<Rectangle[]>;
  proposed_fixes: ProposedFix[];
  formula_fingerprints: Dictionary<Fingerprint>;
  data_fingerprints: Dictionary<Fingerprint>;

  constructor(
    suspicious_cells: ExceLintVector[],
    grouped_formulas: Dictionary<Rectangle[]>,
    grouped_data: Dictionary<Rectangle[]>,
    proposed_fixes: ProposedFix[],
    formula_fingerprints: Dictionary<Fingerprint>,
    data_fingerprints: Dictionary<Fingerprint>
  ) {
    this.suspicious_cells = suspicious_cells;
    this.grouped_formulas = grouped_formulas;
    this.grouped_data = grouped_data;
    this.proposed_fixes = proposed_fixes;
    this.formula_fingerprints = formula_fingerprints;
    this.data_fingerprints = data_fingerprints;
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
  const cmp = vectorComparator(r1.upperleft, r2.upperleft);
  if (cmp == 0) {
    return vectorComparator(r1.bottomright, r2.bottomright);
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
    const formulaCoord = rect.upperleft;
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

  constructor(classification: Classification.BinCategory[], analysis: RectInfo[], direction_is_vert: boolean) {
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
  private wb: WorkbookOutput;

  constructor(wb: WorkbookOutput) {
    this.wb = wb;
  }

  public getSheet(name: string) {
    return this.sheets[name];
  }

  public addSheet(s: WorksheetAnalysis) {
    this.sheets[s.name] = s;
  }

  public get workbook(): WorkbookOutput {
    return this.wb;
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
  get groupedData(): Dictionary<Rectangle[]> {
    return this.analysis.grouped_data;
  }

  // Get the grouped formulas
  get groupedFormulas(): Dictionary<Rectangle[]> {
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
    return this.pf.map((x) => x.score).reduce((x, y) => x + y, 0);
  }

  // Get the total number of anomalous cells
  get numAnomalousCells(): number {
    return this.foundBugs.length;
  }

  // Get the underlying sheet object
  get worksheet(): WorksheetOutput {
    return this.sheet;
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

// /**
//  * Represents the start and end positions of an edit.
//  */
// export class Range {
//   /**
//    * The range's start position in a formula string.
//    */
//   startpos: number;

//   /**
//    * The range's end position.
//    */
//   endpos: number;
// }

// export class Edit {
//   /**
//    * The starting and ending positions of the edit in the cell.
//    */
//   range: Range;
//   /**
//    * The length of the replacement text.
//    */
//   rangeLength: number;
//   /**
//    * The replacement text.
//    */
//   text: string;
//   /**
//    * The address of the cell where the edit occurred.
//    */
//   addr: Address;

//   constructor(range: Range, text: string, addr: Address) {
//     this.range = range;
//     this.text = text;
//     this.addr = addr;
//   }
// }

/**
 * A generic, comparable array.
 */
export class CArray<V extends IComparable<V>> extends Array<V> implements IComparable<CArray<V>> {
  private data: V[];
  constructor(arr: V[]) {
    super();
    this.data = arr;
  }
  public equals(arr: CArray<V>): boolean {
    if (this.data.length != arr.data.length) {
      return false;
    }
    for (let i = 0; i < this.data.length; i++) {
      if (!this.data[i].equals(arr.data[i])) {
        return false;
      }
    }
    return true;
  }

  /**
   * Returns a new CArray formed by concatenating this CArray with
   * the given CArray.  Does not modify given CArrays.
   * @param arr A CArray.
   */
  public concat(arr: CArray<V>): CArray<V> {
    return new CArray(this.data.concat(arr.data));
  }

  public toString(): string {
    return this.data.toString();
  }

  public valueAt(index: number): V {
    return this.data[index];
  }

  public get size(): number {
    return this.data.length;
  }
}

export class Tuple2<T extends IComparable<T>, U extends IComparable<U>> implements IComparable<Tuple2<T, U>> {
  private _elem1: T;
  private _elem2: U;

  constructor(elem1: T, elem2: U) {
    this._elem1 = elem1;
    this._elem2 = elem2;
  }

  public get first(): T {
    return this._elem1;
  }

  public get second(): U {
    return this._elem2;
  }

  public equals(t: Tuple2<T, U>): boolean {
    return this._elem1.equals(t._elem1) && this._elem2.equals(t._elem2);
  }
}

export class Adjacency {
  private _up: Tuple2<Option<Rectangle>, Fingerprint>;
  private _down: Tuple2<Option<Rectangle>, Fingerprint>;
  private _left: Tuple2<Option<Rectangle>, Fingerprint>;
  private _right: Tuple2<Option<Rectangle>, Fingerprint>;

  constructor(
    up: Tuple2<Option<Rectangle>, Fingerprint>,
    down: Tuple2<Option<Rectangle>, Fingerprint>,
    left: Tuple2<Option<Rectangle>, Fingerprint>,
    right: Tuple2<Option<Rectangle>, Fingerprint>
  ) {
    this._up = up;
    this._down = down;
    this._left = left;
    this._right = right;
  }

  public get up() {
    return this._up;
  }

  public get down() {
    return this._down;
  }

  public get left() {
    return this._left;
  }

  public get right() {
    return this._right;
  }
}

export class IncrementalWorkbookAnalysis {}
