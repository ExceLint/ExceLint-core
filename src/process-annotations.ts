'use strict';
let fs = require('fs');

const fname = 'annotations.json';
const content = fs.readFileSync(fname);
let inp = JSON.parse(content);
let out = {}

for (let i = 0; i < inp.length; i++) {
    const workbookName = inp[i]["Workbook"];
    if (!(workbookName in out)) {
	out[workbookName] = {
	    "worksheets" : []
	};
    }
    const sheetName = inp[i]["Worksheet"];
    if (!(sheetName in out[workbookName]["worksheets"])) {
	out[workbookName]["worksheets"].append({
	    "sheet" : sheetName,
	    "bugs" : []
	});
    }
    const bug = {
	"address" : inp[i]["Address"],
	"kind" : inp[i]["BugKind"]
    };
    out[workbookName]["worksheets"][sheetName]["bugs"].push(bug);
}

console.log(JSON.stringify(out));
