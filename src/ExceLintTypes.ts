export interface Dict<V> {
  [key: string]: V;
}

export type Spreadsheet = Array<Array<string>>;

export type DZH = string; // this is for uses of DistinguishedZeroHash

export type ProposedFixes = Array<
  [number, [ExcelintVector, ExcelintVector], [ExcelintVector, ExcelintVector]]
>;

export class ExcelintVector {
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

  public static Zero(): ExcelintVector {
    return new ExcelintVector(0, 0, 0);
  }

  // Subtract other from this vector
  public subtract(v: ExcelintVector): ExcelintVector {
    return new ExcelintVector(this.x - v.x, this.y - v.y, this.c - v.c);
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
}
