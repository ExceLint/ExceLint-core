import { binsearch, strict_binsearch } from "./binsearch";
import { Colorize } from "./colorize";
import {
  ExceLintVector,
  Dict,
  ProposedFix,
  Rectangle,
  Fingerprint,
  Metric,
  upperleft,
  bottomright,
} from "./ExceLintTypes";
import { ExcelJSON } from "./exceljson";

// A comparison function to sort by x-coordinate.
function sort_x_coord(a: Rectangle, b: Rectangle): number {
  const [a1, a2] = a;
  const [b1, b2] = b;
  if (a1.x !== b1.x) {
    return a1.x - b1.x;
  } else {
    return a1.y - b1.y;
  }
}

// A comparison function to sort by y-coordinate.
function sort_y_coord(a: Rectangle, b: Rectangle): number {
  const [a1, a2] = a;
  const [b1, b2] = b;
  if (a1.y !== b1.y) {
    return a1.y - b1.y;
  } else {
    return a1.x - b1.x;
  }
}

// Returns a dictionary containing a bounding box for each group (indexed by hash).
function generate_bounding_box(g: Dict<Rectangle[]>): Dict<Rectangle> {
  const bb: Dict<Rectangle> = {};
  for (const hash of Object.keys(g)) {
    //	console.log("length of formulas for " + i + " = " + g[i].length);
    let xMin = 1000000;
    let yMin = 1000000;
    let xMax = -1000000;
    let yMax = -1000000;

    // find the max/min x and y that bound all the rectangles in the group
    for (let j = 0; j < g[hash].length; j++) {
      const x_tl = g[hash][j][0].x; // top left x
      const x_br = g[hash][j][1].x; // bottom right x
      const y_tl = g[hash][j][0].y; // top left y
      const y_br = g[hash][j][1].y; // bottom right y
      if (x_br > xMax) {
        xMax = x_br;
      }
      if (x_tl < xMin) {
        xMin = x_tl;
      }
      if (y_br > yMax) {
        yMax = y_br;
      }
      if (y_tl < yMin) {
        yMin = y_tl;
      }
    }
    bb[hash] = [new ExceLintVector(xMin, yMin, 0), new ExceLintVector(xMax, yMax, 0)];
    //	console.log("bounding rectangle = (" + xMin + ", " + yMin + "), (" + xMax + ", " + yMax + ")");
  }
  return bb;
}

// Sort formulas in each group by x coordinate
function sort_grouped_formulas(grouped_formulas: Dict<Rectangle[]>): Dict<Rectangle[]> {
  const newGnum: Dict<Rectangle[]> = {};
  for (const key of Object.keys(grouped_formulas)) {
    newGnum[key] = grouped_formulas[key].sort(sort_x_coord);
  }
  return newGnum;
}

// Knuth-Fisher-Yates shuffle (not currently used).
function shuffle<T>(a: Array<T>): Array<T> {
  let j, x, i;
  for (i = a.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    x = a[i];
    a[i] = a[j];
    a[j] = x;
  }
  return a;
}

//test_binsearch();

let comparisons = 0;

function numComparator(a_val: ExceLintVector, b_val: ExceLintVector): number {
  if (a_val.x < b_val.x) {
    return -1;
  }
  if (a_val.x > b_val.x) {
    return 1;
  }
  if (a_val.y < b_val.y) {
    return -1;
  }
  if (a_val.y > b_val.y) {
    return 1;
  }
  if (a_val.c < b_val.c) {
    return -1;
  }
  if (a_val.c > b_val.c) {
    return 1;
  }
  return 0; // they're the same
}

// Return the set of adjacent rectangles that are merge-compatible with the given rectangle
function matching_rectangles(
  rect: Rectangle,
  rect_uls: Array<ExceLintVector>,
  rect_lrs: Array<ExceLintVector>
): Rectangle[] {
  // Assumes uls and lrs are already sorted and the same length.
  const rect_ul = rect[0];
  const rect_lr = rect[1];
  const x1 = rect_ul.x;
  const y1 = rect_ul.y;
  const x2 = rect_lr.x;
  const y2 = rect_lr.y;

  // Try to find something adjacent to A = [[x1, y1, 0], [x2, y2, 0]]
  // options are:
  //   [x1-1, y2] left (lower-right)   [ ] [A] --> [ (?, y1) ... (x1-1, y2) ]
  //   [x2, y1-1] up (lower-right)     [ ]
  //                                   [A] --> [ (x1, ?) ... (x2, y1-1) ]
  //   [x2+1, y1] right (upper-left)   [A] [ ] --> [ (x2 + 1, y1) ... (?, y2) ]
  //   [x1, y2+1] down (upper-left)    [A]
  //                                   [ ] --> [ (x1, y2+1) ... (x2, ?) ]

  // left (lr) = ul_x, lr_y
  const left = new ExceLintVector(x1 - 1, y2, 0);
  // up (lr) = lr_x, ul_y
  const up = new ExceLintVector(x2, y1 - 1, 0);
  // right (ul) = lr_x, ul_y
  const right = new ExceLintVector(x2 + 1, y1, 0);
  // down (ul) = ul_x, lr_y
  const down = new ExceLintVector(x1, y2 + 1, 0);
  const matches: Rectangle[] = [];
  let ind = -1;
  ind = strict_binsearch(rect_lrs, left, numComparator);
  if (ind !== -1) {
    if (rect_uls[ind].y === y1) {
      const candidate: Rectangle = [rect_uls[ind], rect_lrs[ind]];
      matches.push(candidate);
    }
  }
  ind = strict_binsearch(rect_lrs, up, numComparator);
  if (ind !== -1) {
    if (rect_uls[ind].x === x1) {
      const candidate: Rectangle = [rect_uls[ind], rect_lrs[ind]];
      matches.push(candidate);
    }
  }
  ind = strict_binsearch(rect_uls, right, numComparator);
  if (ind !== -1) {
    if (rect_lrs[ind].y === y2) {
      const candidate: Rectangle = [rect_uls[ind], rect_lrs[ind]];
      matches.push(candidate);
    }
  }
  ind = strict_binsearch(rect_uls, down, numComparator);
  if (ind !== -1) {
    if (rect_lrs[ind].x === x2) {
      const candidate: Rectangle = [rect_uls[ind], rect_lrs[ind]];
      matches.push(candidate);
    }
  }
  return matches;
}

let rectangles_count = 0;

// find all merge-compatible rectangles for the given rectangle including their
// fix metrics.
function find_all_matching_rectangles(
  thisfp: string,
  rect: Rectangle,
  fingerprintsX: string[],
  fingerprintsY: string[],
  x_ul: Dict<ExceLintVector[]>,
  x_lr: Dict<ExceLintVector[]>,
  bb: Dict<Rectangle>,
  bbsX: Rectangle[],
  bbsY: Rectangle[]
): ProposedFix[] {
  // get the upper-left and lower-right vectors for the given rectangle
  const [base_ul, base_lr] = rect;

  // this is the output
  let match_list: ProposedFix[] = [];

  // find the index of the given rectangle in the list of rects sorted by X
  const ind1 = binsearch(bbsX, rect, (a: Rectangle, b: Rectangle) => a[0].x - b[0].x);

  // find the index of the given rectangle in the list of rects sorted by Y
  const ind2 = binsearch(bbsY, rect, (a: Rectangle, b: Rectangle) => a[0].y - b[0].y);

  // Pick the coordinate axis that takes us the furthest in the fingerprint list.
  const [fps, itmp, axis] = ind1 > ind2 ? [fingerprintsX, ind1, 0] : [fingerprintsY, ind2, 1];
  const ind = itmp > 0 ? itmp - 1 : itmp;
  for (let i = ind; i < fps.length; i++) {
    const fp = fps[i];
    if (fp === thisfp) {
      continue;
    }
    rectangles_count++;
    // Check bounding box.
    const box = bb[fp];

    /* Since fingerprints are sorted in x-axis order,
	     we can stop once we have gone too far on the x-axis to ever merge again;
	     mutatis mutandis for the y-axis. */

    // early stopping
    if (axis === 0) {
      /* [rect] ... [box]  */
      // if left side of box is too far away from right-most edge of the rectangle
      if (base_lr.x + 1 < box[0].x) {
        break;
      }
    } else {
      /* [rect]
                           ...
                   [box]  */
      // if the top side of box is too far away from bottom-most edge of the rectangle
      if (base_lr.y + 1 < box[0].y) {
        break;
      }
    }

    /*

	      Don't bother processing any rectangle whose edges are
	      outside the bounding box, since they could never be merged with any
	      rectangle inside that box.


                          [ lr_y + 1 < min_y ]

                          +--------------+
      [lr_x + 1 < min_x ] |   Bounding   |  [ max_x + 1 < ul_x ]
	                        |      Box     |
	                        +--------------+

		                      [ max_y + 1 < ul_y ]

	  */

    if (
      base_lr.x + 1 < box[0].x || // left
      base_lr.y + 1 < box[0].y || // top
      box[1].x + 1 < base_ul.x || // right
      box[1].y + 1 < base_ul.y
    ) {
      // Skip. Outside the bounding box.
      //		console.log("outside bounding box.");
    } else {
      const matches: Rectangle[] = matching_rectangles([base_ul, base_lr], x_ul[fp], x_lr[fp]);
      if (matches.length > 0) {
        // compute the fix metric for every potential merge and
        // concatenate them into the match_list
        match_list = match_list.concat(
          matches.map((item: Rectangle) => {
            const metric = Colorize.compute_fix_metric(
              parseFloat(thisfp),
              rect,
              parseFloat(fp),
              item
            );
            return new ProposedFix(metric, rect, item);
          })
        );
      }
    }
  }
  return match_list;
}

// Returns an array with all duplicate proposed fixes removed.
function dedup_fixes(pfs: ProposedFix[]): ProposedFix[] {
  // filtered array
  const rv: ProposedFix[] = [];

  // this is pretty brute force
  for (const i in pfs) {
    const my_pf = pfs[i];
    let found = false;
    for (const j in rv) {
      const oth_pf = rv[j];
      if (my_pf.equals(oth_pf)) {
        found = true;
        break; // my_pf is already in the list
      }
    }
    // add to the list if it was never encountered
    if (!found) rv.push(my_pf);
  }

  return rv;
}

export function find_all_proposed_fixes(grouped_formulas: Dict<Rectangle[]>): ProposedFix[] {
  let all_matches: ProposedFix[] = [];
  rectangles_count = 0;

  // sort each group of rectangles by their x coordinates
  const aNum = sort_grouped_formulas(grouped_formulas);

  // extract from rects the upper-left and lower-right vectors into dicts, indexed by hash
  const x_ul: Dict<ExceLintVector[]> = {}; // upper-left
  const x_lr: Dict<ExceLintVector[]> = {}; // lower-right
  for (const fp of Object.keys(grouped_formulas)) {
    x_ul[fp] = aNum[fp].map((rect) => upperleft(rect));
    x_lr[fp] = aNum[fp].map((rect) => bottomright(rect));
  }

  // find the bounding box for each group
  const bb = generate_bounding_box(grouped_formulas);

  // extract fingerprints
  const fingerprintsX: Fingerprint[] = Object.keys(grouped_formulas);

  // sort fingerprints by the x-coordinate of the upper-left corner of their bounding box.
  fingerprintsX.sort((a: Fingerprint, b: Fingerprint) => bb[a][0].x - bb[b][0].x);

  // generate a sorted list of rectangles
  const bbsX: Rectangle[] = fingerprintsX.map((fp) => bb[fp]);

  // extract fingerprints again
  const fingerprintsY = Object.keys(grouped_formulas);

  // sort fingerprints by the x-coordinate of the upper-left corner of their bounding box.
  fingerprintsY.sort((a: Fingerprint, b: Fingerprint) => bb[a][0].y - bb[b][0].y);

  // generate a sorted list of rectangles
  const bbsY: Rectangle[] = fingerprintsY.map((fp) => bb[fp]);

  // for every group
  for (const fp of Object.keys(grouped_formulas)) {
    // and every rectangle in the group
    for (let i = 0; i < aNum[fp].length; i++) {
      // find all matching rectangles and compute their fix scores
      const matches = find_all_matching_rectangles(
        fp,
        aNum[fp][i],
        fingerprintsX,
        fingerprintsY,
        x_ul,
        x_lr,
        bb,
        bbsX,
        bbsY
      );

      // add these matches to the output
      all_matches = all_matches.concat(matches);
    }
  }

  // reorganize proposed fixes so that the rectangle
  // with the lowest column number comes first
  all_matches = all_matches.map((pf: ProposedFix, _1, _2) => {
    const rect1_ul = upperleft(pf.rect1);
    const rect2_ul = upperleft(pf.rect2);
    // swap rect1 and rect2 depending on the outcome of the comparison
    const newpf: ProposedFix =
      numComparator(rect1_ul, rect2_ul) < 0 ? new ProposedFix(pf.score, pf.rect2, pf.rect1) : pf;
    return newpf;
  });

  // remove duplicate entries
  return dedup_fixes(all_matches);
}
