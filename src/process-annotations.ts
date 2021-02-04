'use strict';
let fs = require('fs');

const fname = 'annotations.json';
const content = fs.readFileSync(fname);
let input = JSON.parse(content);
let out = {}

for (let i = 0; i < input.length; i++) {
    const workbookName = input[i]["Workbook"];
    if (!(workbookName in out)) {
        out[workbookName] = {
            "worksheets": []
        };
    }
    const sheetName = input[i]["Worksheet"];
    if (!(sheetName in out[workbookName]["worksheets"])) {
        out[workbookName]["worksheets"].append({
            "sheet": sheetName,
            "bugs": []
        });
    }
    const bug = {
        "address": input[i]["Address"],
        "kind": input[i]["BugKind"]
    };
    out[workbookName]["worksheets"][sheetName]["bugs"].push(bug);
}

console.log(JSON.stringify(out));
