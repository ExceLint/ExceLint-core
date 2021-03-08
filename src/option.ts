export interface IComparable<V> {
  equals(v: IComparable<V>): boolean;
}

export class Some<T> {
  private t: T;
  public hasValue: true = true;

  constructor(t: T) {
    this.t = t;
  }

  public get value(): T {
    return this.t;
  }

  public equals(o: Option<T>): boolean {
    if (o.hasValue) {
      return this.t === o.t;
    }
    return false;
  }
}
class NoneType {
  public hasValue: false = false;

  public equals(o: Option<any>): boolean {
    return !o.hasValue;
  }
}
export const None = new NoneType(); // singleton None

export type Option<T> = Some<T> | NoneType;

// Given a list of elements of type U and a function that maps elements to
// Option<T>, return only elements of type T.  In other words, filter out
// all NoneType elements, and unwrap Some<T> elements.
export function flatMap<U, T>(f: (u: U) => Option<T>, us: U[]): T[] {
  const ts: T[] = [];
  for (const i in us) {
    const u = us[i];
    const t = f(u);
    // only keep element if it evaluated to Some<T>
    if (t.hasValue) {
      ts.push((t as Some<T>).value);
    }
  }
  return ts;
}

export abstract class Maybe<T extends IComparable<T>> implements IComparable<Maybe<T>> {
  public kind: string = "maybe";
  public abstract equals(o: Maybe<T>): boolean;
}

export class Definitely<T extends IComparable<T>> extends Maybe<T> {
  private t: T;
  public kind: string = "definitely";

  constructor(t: T) {
    super();
    this.t = t;
  }

  public get value(): T {
    return this.t;
  }

  public equals(o: Maybe<T>): boolean {
    if (this.kind === o.kind) {
      return this.t.equals((o as Definitely<T>).t);
    }
    return false;
  }
}

export class Possibly<T extends IComparable<T>> extends Maybe<T> {
  private t: T;
  public kind: string = "possibly";

  constructor(t: T) {
    super();
    this.t = t;
  }

  public get value(): T {
    return this.t;
  }

  public equals(o: Maybe<T>): boolean {
    if (this.kind === o.kind) {
      return this.t.equals((o as Possibly<T>).t);
    }
    return false;
  }
}

export class NoType extends Maybe<any> {
  public kind: string = "no";

  public equals(o: Maybe<any>): boolean {
    return this.kind === o.kind;
  }
}

export const No = new NoType();
