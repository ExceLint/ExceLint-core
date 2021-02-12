/**
 * An implementation of the longest common subsequence algorithm.  Ported from
 * https://github.com/plasma-umass/DataDebug/blob/master/LongestCommonSubsequence/LCS.fs
 *
 * by D. Barowy (2021-02-12)
 */

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

console.log(lcs("hello", "helwordslo"));
