export interface Dict<V> {
  [key: string]: V;
}

export type Spreadsheet = Array<Array<string>>;

export type Fingerprint = string;

export type Metric = number;

// a rectangle is defined by its start and end vectors
export type Rectangle = [ExceLintVector, ExceLintVector];

/* a tuple where:
   - Metric is a "fix distance",
   - the first Rectangle is the suspected buggy rectangle, and
   - the second Rectangle is the merge candidate.
*/
export type ProposedFix = [Metric, Rectangle, Rectangle];

export function rectangles(pf: ProposedFix): Rectangle[] {
  const [_, rect1, rect2] = pf;
  return [rect1, rect2];
}

export function rect1(pf: ProposedFix): Rectangle {
  return pf[1];
}

export function rect2(pf: ProposedFix): Rectangle {
  return pf[2];
}

export function upperleft(r: Rectangle): ExceLintVector {
  return r[0];
}

export function bottomright(r: Rectangle): ExceLintVector {
  return r[1];
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

  // vector sum reduction
  public static readonly VectorSum = (acc: ExceLintVector, curr: ExceLintVector): ExceLintVector =>
    new ExceLintVector(acc.x + curr.x, acc.y + curr.y, acc.c + curr.c);
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
