"use strict";

import path = require("path");
import * as XLSX from "xlsx";
import * as sha224 from "crypto-js/sha224";
import * as base64 from "crypto-js/enc-base64";
import { Dict, Spreadsheet } from "./ExceLintTypes";

export class WorksheetOutput {
  sheetName: string;
  usedRangeAddress: string;
  formulas: Spreadsheet;
  values: Spreadsheet;
  styles: Spreadsheet;

  constructor(
    sheetName: string,
    usedRangeAddress: string,
    formulas: Spreadsheet,
    values: Spreadsheet,
    styles: Spreadsheet
  ) {
    this.sheetName = sheetName;
    this.usedRangeAddress = usedRangeAddress;
    this.formulas = formulas;
    this.values = values;
    this.styles = styles;
  }
}

export class WorkbookOutput {
  workbookName: string;
  worksheets: WorksheetOutput[];

  constructor(filename: string) {
    this.workbookName = filename;
    this.worksheets = [];
  }

  // Tracks a worksheet object in this workbook object
  public addWorksheet(ws: WorksheetOutput): void {
    this.worksheets.push(ws);
  }

  // Returns the filename of the workbook, independently of the path
  public workbookBaseName(): string {
    return path.basename(this.workbookName);
  }

  // Makes a copy of a WorkbookOutput object, replacing the name
  public static AdjustWorkbookName(wb: WorkbookOutput) {
    const wbnew = new WorkbookOutput(wb.workbookBaseName());
    wbnew.worksheets = wb.worksheets;
    return wbnew;
  }
}

export class ExcelJSON {
  private static general_re = "\\$?[A-Z][A-Z]?\\$?\\d+"; // column and row number, optionally with $
  private static pair_re = new RegExp(
    "(" + ExcelJSON.general_re + "):(" + ExcelJSON.general_re + ")"
  );
  private static singleton_re = new RegExp(ExcelJSON.general_re);

  public static processWorksheet(sheet, selection: ExcelJSON.selections) {
    let ref = "A1:A1"; // for empty sheets.
    if ("!ref" in sheet) {
      // Not empty.
      ref = sheet["!ref"];
    }
    const decodedRange = XLSX.utils.decode_range(ref);
    const startColumn = 0; // decodedRange['s']['c'];
    const startRow = 0; // decodedRange['s']['r'];
    const endColumn = decodedRange["e"]["c"];
    const endRow = decodedRange["e"]["r"];

    const rows: string[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: string[] = [];
      for (let c = startColumn; c <= endColumn; c++) {
        const cell = XLSX.utils.encode_cell({ c: c, r: r });
        const cellValue = sheet[cell];
        // console.log(cell + ': ' + JSON.stringify(cellValue));
        let cellValueStr = "";
        if (cellValue) {
          switch (selection) {
            case ExcelJSON.selections.FORMULAS:
              if (!cellValue["f"]) {
                cellValueStr = "";
              } else {
                cellValueStr = "=" + cellValue["f"];
              }
              break;
            case ExcelJSON.selections.VALUES:
              // Numeric values.
              if (cellValue["t"] === "n") {
                if ("z" in cellValue && cellValue["z"] && cellValue["z"].endsWith("yy")) {
                  // ad hoc date matching.
                  // skip dates.
                } else {
                  cellValueStr = JSON.stringify(cellValue["v"]);
                }
              }
              break;
            case ExcelJSON.selections.STYLES:
              if (cellValue["s"]) {
                // Encode the style as a hash (and just keep the first 10 characters).
                const styleString = JSON.stringify(cellValue["s"]);
                const str = base64.stringify(sha224(styleString));
                cellValueStr = str.slice(0, 10);
              }
              break;
          }
        }
        row.push(cellValueStr);
      }
      rows.push(row);
    }
    return rows;
  }

  public static processWorkbookFromXLSX(f: XLSX.WorkBook, filename: string): WorkbookOutput {
    const output = new WorkbookOutput(filename);
    const sheetNames = f.SheetNames;
    const sheets: Dict<XLSX.WorkSheet> = f.Sheets;
    for (const sheetName of sheetNames) {
      const sheet = sheets[sheetName];
      if (!sheet) {
        // Weird edge case here.
        continue;
      }
      // console.warn('  processing ' + filename + '!' + sheetName);
      // Try to parse the ref to see if it's a pair (e.g., A1:B10) or a singleton (e.g., C9).
      // If the latter, make it into a pair (e.g., C9:C9).
      let ref;
      if ("!ref" in sheet) {
        ref = sheet["!ref"];
      } else {
        // Empty sheet.
        ref = "A1:A1";
      }
      const result = ExcelJSON.pair_re.exec(ref); // ExcelJSON.pair_re.exec(ref);
      if (result) {
        // It's a pair; we're fine.
        // ACTUALLY to work around a bug downstream, we start everything at A1.
        // This sucks but it works.
        ref = "A1:" + result[2];
      } else {
        // Singleton. Make it a pair.
        ref = ref + ":" + ref;
      }
      const sheetRange = sheetName + "!" + ref;
      const sheet_formulas = ExcelJSON.processWorksheet(sheet, ExcelJSON.selections.FORMULAS);
      const sheet_values = ExcelJSON.processWorksheet(sheet, ExcelJSON.selections.VALUES);
      const sheet_styles = ExcelJSON.processWorksheet(sheet, ExcelJSON.selections.STYLES);
      const wso = new WorksheetOutput(
        sheetName,
        sheetRange,
        sheet_formulas,
        sheet_values,
        sheet_styles
      );
      output.addWorksheet(wso);
    }
    return output;
  }

  public static processWorkbook(base: string, filename: string): WorkbookOutput {
    const f = XLSX.readFile(base + filename, { cellStyles: true });
    return this.processWorkbookFromXLSX(f, filename);
  }
}

export namespace ExcelJSON {
  export enum selections {
    FORMULAS,
    VALUES,
    STYLES,
  }
}
