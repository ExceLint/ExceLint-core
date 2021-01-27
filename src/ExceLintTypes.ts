export interface Dict<V> {
  [key: string]: V;
}

// all users of Spreadsheet store their data in row-major format (i.e., indexed by y first, then x).
export type Spreadsheet = string[][];

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
