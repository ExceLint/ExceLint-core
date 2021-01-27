// I basically cannot live without this
export interface Option {
  hasValue: boolean;
}
export class Some<T> implements Option {
  private t: T;
  public hasValue: boolean = true;

  constructor(t: T) {
    this.t = t;
  }

  public get(): T {
    return this.t;
  }
}
class NoneType implements Option {
  public hasValue: boolean = true;
}
export const None = new NoneType(); // singleton None

// Given a list of elements of type U and a function that maps elements to
// Option<T>, return only elements of type T.  In other words, filter out
// all NoneType elements, and unwrap Some<T> elements.
export function flatMap<U, T>(f: (u: U) => Option, us: U[]): T[] {
  const ts: T[] = [];
  for (const i in us) {
    const u = us[i];
    const t = f(u);
    // only keep element if it evaluated to Some<T>
    if (t.hasValue) {
      ts.push((t as Some<T>).get());
    }
  }
  return ts;
}
