"use strict";
import { Dictionary } from "./ExceLintTypes";
let fs = require("fs");

const fname = "annotations.json";
const content = fs.readFileSync(fname);
let input = JSON.parse(content);
let out = new Dictionary<any>();

for (let i = 0; i < input.length; i++) {
  const workbookName = input[i]["Workbook"];
  if (!(workbookName in out)) {
    out.put(workbookName, {
      worksheets: [],
    });
  }
  const sheetName = input[i]["Worksheet"];
  if (!(sheetName in out.get(workbookName)["worksheets"])) {
    out.get(workbookName)["worksheets"].append({
      sheet: sheetName,
      bugs: [],
    });
  }
  const bug = {
    address: input[i]["Address"],
    kind: input[i]["BugKind"],
  };
  out.get(workbookName)["worksheets"][sheetName]["bugs"].push(bug);
}

console.log(JSON.stringify(out));
