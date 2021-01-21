export class JSONclone {
  // This method performs a type-safe deep copy of any
  // JavaScript value.
  public static clone<T>(data: T): T {
    if (data) {
      // if data is an array, do an array copy
      if (Array.isArray(data)) {
        // this is a hack to keep things type-safe
        return (data.slice() as unknown) as T;
      }
      // if data is an object, recursively copy fields
      // and assign them to the copied object
      if (data.constructor === Object) {
        const obj = {} as T;
        for (const k of Object.keys(data)) {
          obj[k] = JSONclone.clone(data[k]);
        }
        return obj;
      }
      return data;
    }
    // if data is null, there is no copying to be done
    return null;
  }
}
