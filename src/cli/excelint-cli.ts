// Process Excel files (input from .xls or .xlsx) with ExceLint.
// by Emery Berger, Microsoft Research / University of Massachusetts Amherst
// www.emeryberger.com

"use strict";
import path = require("path");
import { ExcelJSON } from "../exceljson";
import { Colorize } from "../colorize";
import { WorkbookAnalysis } from "../ExceLintTypes";
import { Config } from "../config";
import { CLIConfig, process_arguments } from "./args";
import { AnnotationData } from "./bugs";

const BUG_DATA_PATH = "test/annotations-processed.json";

//
// Process arguments.
//
const args: CLIConfig = process_arguments();

//
// Ready to start processing.
//

// open annotations file
const theBugs = new AnnotationData(BUG_DATA_PATH);

// get base directory
const base = args.directory ? args.directory + "/" : "";

// for each parameter setting, run analyses on all files
const outputs: WorkbookAnalysis[] = [];
for (const parms of args.parameters) {
  const formattingDiscount = parms[0];
  Config.setFormattingDiscount(formattingDiscount);
  const reportingThreshold = parms[1];
  Config.setReportingThreshold(reportingThreshold);

  // process every file given by the user
  for (const fname of args.allFiles) {
    args.numWorkbooks += 1;

    // Open the given input spreadsheet
    console.warn("processing " + fname);
    const inp = ExcelJSON.processWorkbook(base, fname);

    // Find out a few facts about this workbook in the bug database
    const facts = theBugs.check(inp);
    if (facts.hasError) {
      args.numSheetsWithErrors += 1;
      args.numWorkbooksWithErrors += 1;
    }
    if (facts.hasFormula) args.numWorkbooksWithFormulas += 1;
    args.numSheets += facts.numSheets;

    const output = Colorize.process_workbook(inp, ""); // no bug processing for now; just get all sheets
    outputs.push(output);
  }
}

if (!args.suppressOutput) {
  console.log(JSON.stringify(outputs, null, "\t"));
}
