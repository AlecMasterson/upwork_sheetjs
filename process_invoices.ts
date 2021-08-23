import ExcelJS, { Cell, Workbook, Worksheet } from "exceljs";
import { COLUMNS_INVOICE, COLUMNS_RETURNS } from "./configs/ColumnConfig";
import Batch from "./types/Batch";
import { ExcelColumn } from "./types/ExcelColumns";
import { add_formatted_table, adjust_column_widths, exportWorkbook } from "./utils/Util";
import * as FileStream from "fs";
import moment from "moment";

/* =================================================
====================================================
CONFIGURATION AND PREPARATION
====================================================
================================================= */

// The JSON input file.
const INPUT_FILE_NAME = "data/example.json";

/* =================================================
====================================================
FUNCTIONS
====================================================
================================================= */

// Function to process the batches of an account.
// This function will create and export the XLSX file associated with that account.
function processAccount(batches: Batch[]): void {
    const workbook: Workbook = new ExcelJS.Workbook();
    const sheetInvoice: Worksheet = workbook.addWorksheet("invoice");
    const sheetReturnsAndOther: Worksheet = workbook.addWorksheet("returns_and_other");

    // Add the "Invoice Due" data with today's date.
    const parsedBatches: Batch[] = batches.map((batch: Batch) => ({ ...batch, "Invoice Due": new Date() }));

    // Add and format the data as a table in the Worksheet.
    add_formatted_table(sheetInvoice, "invoice_table", COLUMNS_INVOICE, parsedBatches, true);

    // Add and format the headers to the blank Worksheet.
    sheetReturnsAndOther.columns = COLUMNS_RETURNS.map((column: ExcelColumn) => ({ header: column.name }));
    sheetReturnsAndOther.getRow(1).font = { bold: true };

    // Add the formula for the "Returns & Other" column. Exclude the header and footer rows.
    sheetInvoice.getColumn("F").eachCell((cell: Cell, index: number) => {
        if (index === 1 || index === parsedBatches.length + 2) return;
        cell.value = { date1904: false, formula: `SUMIFS(returns_and_other!C:C, returns_and_other!A:A, B${index})` };
    })

    // Add formulas for the columns "Total", "If Paid By Credit Card", and "Invoice Due" respectively.
    sheetInvoice.fillFormula(`G2:G${parsedBatches.length + 1}`, "D2+E2+F2");
    sheetInvoice.fillFormula(`H2:H${parsedBatches.length + 1}`, "G2*1.03");
    sheetInvoice.fillFormula(`J2:J${parsedBatches.length + 1}`, "I2+7");

    // Fit each column width to the maximum cell length for that column plus a buffer.
    adjust_column_widths(sheetReturnsAndOther, COLUMNS_RETURNS);

    // Export the Workbook to an XLSX file.
    const date: string = moment().format("MM-DD-yyyy");
    exportWorkbook(workbook, `reports/${date}/invoices`, `${parsedBatches[0].Account}_invoice_${date}`);
}

/* =================================================
====================================================
MAIN PROCESS
====================================================
================================================= */

// Read the JSON input data and process each account separately.
const data: { [key: string]: Batch[] } = JSON.parse(FileStream.readFileSync(INPUT_FILE_NAME).toString());
Object.keys(data).forEach((accountNumber: string) => processAccount(data[accountNumber]));
