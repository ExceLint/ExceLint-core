// Process Excel files (input from .xls or .xlsx) with ExceLint.
// by Emery Berger, Microsoft Research / University of Massachusetts Amherst
// www.emeryberger.com

"use strict";
import { ExcelJSON } from "../exceljson";
import { Colorize } from "../colorize";
import { Dictionary, WorkbookAnalysis } from "../ExceLintTypes";
import { Config } from "../config";
import { CLIConfig, process_arguments } from "./args";
import { AnnotationData } from "./bugs";
import { Timer } from "../timer";

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
// an array of times indexed by benchmark name
const times = new Dictionary<number[]>();
for (const parms of args.parameters) {
  const formattingDiscount = parms[0];
  Config.setFormattingDiscount(formattingDiscount);
  const reportingThreshold = parms[1];
  Config.setReportingThreshold(reportingThreshold);

  // process every file given by the user
  for (const fname of args.allFiles) {
    args.numWorkbooks += 1;

    console.warn("processing " + fname);

    // initialize array before first run
    const time_arr: number[] = [];
    times.put(fname, time_arr);

    for (let i = 0; i < args.runs; i++) {
      // Open the given input spreadsheet
      const inp = ExcelJSON.processWorkbook(base, fname);

      // Find out a few facts about this workbook in the bug database
      const facts = theBugs.check(inp);
      if (facts.hasError) {
        args.numSheetsWithErrors += 1;
        args.numWorkbooksWithErrors += 1;
      }
      if (facts.hasFormula) args.numWorkbooksWithFormulas += 1;
      args.numSheets += facts.numSheets;

      const t = new Timer("full analysis");
      const output = Colorize.process_workbook(inp, "", !args.suppressOutput); // no bug processing for now; just get all sheets
      const elapsed_us = t.elapsedTime();
      outputs.push(output);
      time_arr.push(elapsed_us);
    }
  }
}

if (!args.suppressOutput && !args.elapsedTime) {
  console.log(JSON.stringify(outputs, null, "\t"));
}

if (args.elapsedTime) {
  console.log("Full analysis times:");
  const workbooks = times.keys;
  console.log('"workbook","trial","time_μs"');
  for (let i = 0; i < workbooks.length; i++) {
    const workbook = workbooks[i];
    const wTimes = times.get(workbook);
    for (let j = 0; j < wTimes.length; j++) {
      console.log('"' + workbook + '","' + j + '","' + Timer.round(wTimes[j]) + '"');
    }
  }
}
