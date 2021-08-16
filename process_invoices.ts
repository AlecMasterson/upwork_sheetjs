import ExcelJS, { Cell, Workbook, Worksheet } from "exceljs";
import { COLUMNS_INVOICE, COLUMNS_RETURNS } from "./configs/ColumnConfig";
import Batch from "./types/Batch";
import { ExcelColumn } from "./types/Columns";
import { adjust_column_widths, exportWorkbook, parseExcelColumn } from "./utils/Util";
import * as FileStream from "fs";
import Lodash from "lodash";
import moment from "moment";

// Read the JSON input data and process each account separately.
const data: { [key: string]: Batch[] } = JSON.parse(FileStream.readFileSync(process.argv[2]).toString());
Object.keys(data).forEach((accountNumber: string) => processAccount(data[accountNumber]));

function processAccount(batches: Batch[]): void {
    const workbook: Workbook = new ExcelJS.Workbook();
    const sheetInvoice: Worksheet = workbook.addWorksheet("invoice");
    const sheetReturnsAndOther: Worksheet = workbook.addWorksheet("returns_and_other");

    // Add the "Invoice Due" data with today's date.
    const parsedBatches: Batch[] = batches.map((batch: Batch) => ({ ...batch, "Invoice Due": new Date() }));

    sheetInvoice.addTable({
        columns: COLUMNS_INVOICE.map((column: ExcelColumn) => parseExcelColumn(column)),
        name: "invoice_table",
        ref: "A1",
        rows: parsedBatches.map((batch: Batch) => COLUMNS_INVOICE.map((column: ExcelColumn) => batch[column.key])),
        style: { showRowStripes: true },
        totalsRow: true
    });

    // Add the blank sheet and apply formatting, bold the header row.
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

    // Apply cell formatting to each column, excluding the header row.
    COLUMNS_INVOICE.forEach((column: ExcelColumn, i: number) => sheetInvoice.getColumn(i + 1).numFmt = Lodash.get(column, "style.numFmt", null));
    sheetInvoice.getRow(1).numFmt = null;

    // Apply column formatting, fit the width to the cell data.
    adjust_column_widths(sheetInvoice, COLUMNS_INVOICE);
    adjust_column_widths(sheetReturnsAndOther, COLUMNS_RETURNS);

    // Export the Workbook to an XLSX file.
    const date: string = moment().format("MM-DD-yyyy");
    exportWorkbook(workbook, `reports/${date}/invoices`, `${parsedBatches[0].Account}_invoice_${date}`);
}
