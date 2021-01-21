export interface Dict<V> {
  [key: string]: V;
}

export type Spreadsheet = Array<Array<string>>;

export type Fingerprint = string;

export type Metric = number;

// a rectangle is defined by its start and end vectors
export type Rectangle = [ExceLintVector, ExceLintVector];

export type ProposedFixes = [number, Rectangle][];

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
