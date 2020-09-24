export class JSONclone {
  public static clone(data: any): any {
    if (data) {
      if (Array.isArray(data)) {
        return data.slice();
      }
      if (data.constructor === Object) {
        const obj = {};
        for (const k of Object.keys(data)) {
          obj[k] = JSONclone.clone(data[k]);
        }
        return obj;
      }
      return data;
    }
    return null;
  }
}
