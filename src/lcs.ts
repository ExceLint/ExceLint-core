/**
 * An implementation of the longest common subsequence algorithm.  Ported from
 * https://github.com/plasma-umass/DataDebug/blob/master/LongestCommonSubsequence/LCS.fs
 *
 * by D. Barowy (2021-02-12)
 */

import { IComparable, CSet } from "./ExceLintTypes";

class NumPair implements IComparable<NumPair> {
  private fst: number;
  private snd: number;

  public equals(v: NumPair): boolean {
    return this.first === v.first && this.second === v.second;
  }

  public get first(): number {
    return this.fst;
  }

  public get second(): number {
    return this.second;
  }
}

/**
 * Initialize a 2D array and fill it with a value;
 * @param value A value of type T.
 * @param m The size of the first array dimension.
 * @param n The size of the second array dimension.
 */
export function fill2D<T>(value: T, m: number, n: number): T[][] {
  const arr: T[][] = [];
  for (let i = 0; i < m; i++) {
    arr[i] = [];
    for (let j = 0; j < n; j++) {
      arr[i][j] = value;
    }
  }
  return arr;
}

/**
 * Computes the set of longest strings.
 * @param x One string.
 * @param y Another string.
 */
export function lcs(x: string, y: string): string[] {
  const m = x.length;
  const n = y.length;
  const C = lcsLength(x, m, y, n);
  return backtrackAll(C, x, m, y, n);
}

/**
 * Returns a dynamic programming table of longest matches between x and y.
 * @param x String x.
 * @param m The length of string x.
 * @param y String y.
 * @param n The length of string y.
 */
function lcsLength(x: string, m: number, y: string, n: number): number[][] {
  const C = fill2D(0, m + 1, n + 1);
  for (let i = 1; i < m + 1; i++) {
    for (let j = 1; j < n + 1; j++) {
      console.log("Processing: i = " + i.toString() + ", j = " + j.toString());
      // are the characters the same at this position?
      if (x.charAt(i - 1) === y.charAt(j - 1)) {
        // then the length of this edit is the length of
        // the previous edits up to this point, plus one.
        C[i][j] = C[i - 1][j - 1] + 1;
      } else {
        // otherwise, the length is the longest edit of
        // either x or y
        C[i][j] = Math.max(C[i][j - 1], C[i - 1][j]);
      }
    }
  }
  return C;
}

/**
 * Computes the set union of sets a and b, returning an array of strings.
 * @param a Set of strings a.
 * @param b Set of strings b.
 */
function union(a: string[], b: string[]): string[] {
  const uniq = new Set<string>();
  for (let i = 0; i < a.length; i++) {
    uniq.add(a[i]);
  }
  for (let i = 0; i < a.length; i++) {
    uniq.add(b[i]);
  }
  return Array.from(uniq);
}

/**
 * Returns the set of all longest subsequences.
 * @param C A dynamic programming table representing matches between x and y.
 * @param x String x.
 * @param i Length of string x.
 * @param y String y.
 * @param j Length of string y.
 */
function backtrackAll(C: number[][], x: string, i: number, y: string, j: number): string[] {
  if (i === 0 || j === 0) {
    // if both indices are zero, we're just starting
    return [""];
  } else if (x.charAt(i - 1) === y.charAt(j - 1)) {
    // otherwise, if the characters are the same at this position,
    // backtrack and append the matching character to the end of
    // each string in the set.
    const Z = backtrackAll(C, x, i - 1, y, j - 1);
    return Z.map((z: string) => z + x.charAt(i - 1));
  } else {
    // if they're not the same...
    let R: string[] = [];
    // find which subsequence is the longest
    // note: both possibilities can be the longest
    if (C[i][j - 1] >= C[i - 1][j]) {
      // if C[i][j-1] is the longer subsequence
      R = backtrackAll(C, x, i, y, j - 1);
    }
    if (C[i - 1][j] >= C[i][j - 1]) {
      // if C[i-1][j] is the longer subsequence
      R = union(R, backtrackAll(C, x, i - 1, y, j));
    }
    return R;
  }
}

/**
 * // like backtrackAll except that it returns a set of character pair
    // sequences instead of a set of strings
    // for each character pair: (X pos, Y pos)
    let rec getCharPairs(C: int[,], X: string, Y: string, i: int, j: int, sw: Stopwatch, timeout: TimeSpan) : Set<(int*int) list> =
        if sw.Elapsed > timeout then
            raise (TimeoutException())
        if i = 0 || j = 0 then
            set[[]]
        else if X.[i-1] = Y.[j-1] then
            let mutable ZS = Set.map (fun (Z: (int*int) list) -> Z @ [(i-1,j-1)] ) (getCharPairs(C, X, Y, i-1, j-1, sw, timeout))
            if (C.[i,j] = C.[i,j-1]) then 
                ZS <- Set.union ZS (getCharPairs(C, X, Y, i, j-1, sw, timeout))
            ZS
        else
            let mutable R = Set.empty
            if C.[i,j-1] >= C.[i-1,j] then
                R <- getCharPairs(C, X, Y, i, j-1, sw, timeout)
            if C.[i-1,j] >= C.[i,j-1] then
                R <- Set.union R (getCharPairs(C, X, Y, i-1, j, sw, timeout))
            R
 */

function getCharPairs(C: number[][], x: string, i: number, y: string, j: number): CSet<CSet<NumPair>> {
  if (i === 0 || j === 0) {
    const outer: CSet<CSet<NumPair>> = CSet.empty();
    outer.add(CSet.empty());
    return outer;
  } else if (x.charAt(i - 1) === y.charAt(j - 1)) {
    const Z = getCharPairs(C, x, i - 1, y, j - 1);
    let ZS = Z.map((z: Pair<number, number>[]) => z.concat([[i - 1, j - 1]]));
    if (C[i][j] === C[i][j - 1]) {
      const W = getCharPairs(C, x, i, y, j - 1);
      ZS = union(ZS, W);
    }
  }
  return [];
}

console.log(lcs("hello", "helwordslo"));
