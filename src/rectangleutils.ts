import { ExceLintVector, Rectangle } from "./ExceLintTypes";

export class RectangleUtils {
  public static is_adjacent(A: Rectangle, B: Rectangle): boolean {
    const [a1, a2] = A;
    const [b1, b2] = B;

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
    const [a1, a2] = A;
    const [b1, b2] = B;
    return [
      new ExceLintVector(Math.min(a1.x, b1.x), Math.min(a1.y, b1.y), 0),
      new ExceLintVector(Math.max(a2.x, b2.x), Math.max(a2.y, b2.y), 0),
    ];
  }

  public static area(A: Rectangle): number {
    const [a1, a2] = A;
    const length = a2.x - a1.x + 1;
    const width = a2.y - a1.y + 1;
    return length * width;
  }

  public static diagonal(A: Rectangle): number {
    const [a1, a2] = A;
    const length = a2.x - a1.x + 1;
    const width = a2.y - a1.y + 1;
    return Math.sqrt(length * length + width * width);
  }

  public static overlap(A: Rectangle, B: Rectangle): number {
    const [a1, a2] = A;
    const [b1, b2] = B;
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
