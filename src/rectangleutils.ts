import { ExceLintVector, Rectangle, Range, Address } from "./ExceLintTypes";

export class FatCross {
  public readonly up: Range;
  public readonly down: Range;
  public readonly left: Range;
  public readonly right: Range;

  constructor(up: Range, down: Range, left: Range, right: Range) {
    this.up = up;
    this.down = down;
    this.left = left;
    this.right = right;
  }
}

export class RectangleUtils {
  public static is_adjacent(A: Rectangle, B: Rectangle): boolean {
    const a1 = A.upperleft;
    const a2 = A.bottomright;
    const b1 = B.upperleft;
    const b2 = B.bottomright;

    const tolerance = 1;
    const adj = !(
      a1.x - b2.x > tolerance ||
      b1.x - a2.x > tolerance ||
      a1.y - b2.y > tolerance ||
      b1.y - a2.y > tolerance
    );
    return adj;
  }

  public static bounding_box(A: Rectangle, B: Rectangle): Rectangle {
    const a1 = A.upperleft;
    const a2 = A.bottomright;
    const b1 = B.upperleft;
    const b2 = B.bottomright;

    return new Rectangle(
      new ExceLintVector(Math.min(a1.x, b1.x), Math.min(a1.y, b1.y), 0),
      new ExceLintVector(Math.max(a2.x, b2.x), Math.max(a2.y, b2.y), 0)
    );
  }

  public static area(A: Rectangle): number {
    const a1 = A.upperleft;
    const a2 = A.bottomright;
    const length = a2.x - a1.x + 1;
    const width = a2.y - a1.y + 1;
    return length * width;
  }

  public static diagonal(A: Rectangle): number {
    const a1 = A.upperleft;
    const a2 = A.bottomright;
    const length = a2.x - a1.x + 1;
    const width = a2.y - a1.y + 1;
    return Math.sqrt(length * length + width * width);
  }

  public static overlap(A: Rectangle, B: Rectangle): number {
    const a1 = A.upperleft;
    const a2 = A.bottomright;
    const b1 = B.upperleft;
    const b2 = B.bottomright;
    let width = 0,
      height = 0;
    if (a2.x > b2.x) {
      width = b2.x - a1.x + 1;
    } else {
      width = a2.x - b1.x + 1;
    }
    if (a2.y > b2.y) {
      height = b2.y - a1.y + 1;
    } else {
      height = a2.y - b1.y + 1;
    }
    return width * height; // Math.max(0, Math.min(ax2, bx2) - Math.max(ax1, bx1)) * Math.max(0, Math.min(ay2, by2) - Math.max(ay1, by1));
  }

  public static is_mergeable(A: Rectangle, B: Rectangle): boolean {
    return (
      RectangleUtils.is_adjacent(A, B) &&
      RectangleUtils.area(A) + RectangleUtils.area(B) - RectangleUtils.overlap(A, B) ===
        RectangleUtils.area(RectangleUtils.bounding_box(A, B))
    );
  }

  /**
   * Returns true if the target is contained within the given range.
   * @param rng A range.
   * @param target The cell in question.
   * @returns True if the target is inside the range.
   */
  public static targetInRange(rng: Range, target: Address): boolean {
    return (
      target.column >= rng.upperLeftColumn &&
      target.column <= rng.bottomRightColumn &&
      target.row >= rng.upperLeftRow &&
      target.row <= rng.bottomRightRow
    );
  }

  /**
   * Ensure that the target range is entirely contained within the containing range.
   * @param rngContainer Containing range.
   * @param rngTarget Target range
   * @returns A truncated Range.
   */
  public static truncateRangeInRange(rngContainer: Range, rngTarget: Range): Range {
    if (rngContainer.addressStart.worksheet !== rngTarget.addressStart.worksheet) {
      throw new Error("Both ranges must be on the same worksheet.");
    }
    const sheet = rngContainer.addressStart.worksheet;
    const leftmost_x = Math.max(rngContainer.upperLeftColumn, rngTarget.upperLeftColumn);
    const rightmost_x = Math.min(rngContainer.bottomRightColumn, rngTarget.bottomRightColumn);
    const uppermost_y = Math.max(rngContainer.upperLeftRow, rngTarget.upperLeftRow);
    const bottommost_y = Math.min(rngContainer.bottomRightRow, rngTarget.bottomRightRow);
    return new Range(new Address(sheet, uppermost_y, leftmost_x), new Address(sheet, bottommost_y, rightmost_x));
  }

  /**
   * Finds the dimensions of the four analysis regions relevant to the given
   * target inside the given range. Note that regions are not likely to be
   * the same size.
   * @param rng A rectangular region.
   * @param target The cell being analyzed.
   */
  public static findFatCross(rng: Range, target: Address): FatCross {
    // sanity check
    RectangleUtils.targetInRange(rng, target);

    // get sheet
    const sheet = rng.addressStart.worksheet;

    // top region
    const up_ul = new Address(sheet, rng.upperLeftRow, target.column - 1);
    const up_br = new Address(sheet, target.row, target.column + 1);
    const up = RectangleUtils.truncateRangeInRange(rng, new Range(up_ul, up_br));

    // bottom region
    const down_ul = new Address(sheet, target.row + 1, target.column - 1);
    const down_br = new Address(sheet, rng.bottomRightRow, target.column + 1);
    const down = RectangleUtils.truncateRangeInRange(rng, new Range(down_ul, down_br));

    // left region
    const left_ul = new Address(sheet, target.row - 1, rng.upperLeftColumn);
    const left_br = new Address(sheet, target.row + 1, target.column - 2);
    const left = RectangleUtils.truncateRangeInRange(rng, new Range(left_ul, left_br));

    // right region
    const right_ul = new Address(sheet, target.row - 1, target.column + 2);
    const right_br = new Address(sheet, target.row + 1, rng.bottomRightColumn);
    const right = RectangleUtils.truncateRangeInRange(rng, new Range(right_ul, right_br));

    return new FatCross(up, down, left, right);
  }

  /*
        public static testme() {
        console.assert(RectangleUtils.is_mergeable([ [ 1, 1 ], [ 1, 1 ] ], [ [ 2, 1 ], [ 2, 1 ] ]), "nope1");
        console.assert(RectangleUtils.is_mergeable([ [ 1, 1 ], [ 1, 10 ] ], [ [ 2, 1 ], [ 2, 10 ] ]), "nope2");
        console.assert(RectangleUtils.is_mergeable([ [ 2, 2 ], [ 4, 4 ] ], [ [ 5, 2 ], [ 8, 4 ] ]), "nope3");
        console.assert(!RectangleUtils.is_mergeable([ [ 2, 2 ], [ 4, 4 ] ], [ [ 4, 2 ], [ 8, 5 ] ]), "nope4");
        console.assert(!RectangleUtils.is_mergeable([ [ 1, 1 ], [ 1, 10 ] ], [ [ 2, 1 ], [ 2, 11 ] ]), "nope5");
        console.assert(!RectangleUtils.is_mergeable([ [ 1, 1 ], [ 1, 10 ] ], [ [ 3, 1 ], [ 3, 10 ] ]), "nope6");
        console.assert(RectangleUtils.is_mergeable([ [ 2, 7 ], [ 3, 11 ] ], [ [ 3, 7 ], [ 4, 11 ] ]), "nope7");
        }
    */
}
