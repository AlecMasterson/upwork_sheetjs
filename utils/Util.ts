import { Cell, Column, Workbook, Worksheet } from "exceljs";
import { ExcelColumn, MIN_COLUMN_WIDTH } from "../types/ExcelColumns";
import * as FileStream from "fs";
import Lodash from "lodash";

export function add_formatted_table(sheet: Worksheet, name: string, columns: ExcelColumn[], data: object[], totalsRow = false): void {
    sheet.addTable({
        columns: columns.map((column: ExcelColumn) => ({
            name: column.name,
            filterButton: column.filterButton,
            totalsRowFunction: column.totalsRowFunction,
            totalsRowLabel: column.totalsRowLabel
        })),
        name,
        ref: "A1",
        rows: data.map((row: object) => columns.map((column: ExcelColumn) => row[column.key])),
        style: { showRowStripes: true },
        totalsRow
    });

    // Apply cell formatting to each column, excluding the header row.
    columns.forEach((column: ExcelColumn, i: number) => sheet.getColumn(i + 1).numFmt = Lodash.get(column, "style.numFmt", null));
    sheet.getRow(1).numFmt = null;

    // Fit each column width to the maximum cell length for that column plus a buffer.
    adjust_column_widths(sheet, columns);
}

// Function for adjusting each column in the Worksheet to fit the data in that column.
export function adjust_column_widths(sheet: Worksheet, columns: ExcelColumn[]): void {
    sheet.columns.forEach((column: Partial<Column>, i: number) => {
        let minLength: number = MIN_COLUMN_WIDTH;

        // Get the longest length of content in a cell for the given column.
        column.eachCell && column.eachCell({ includeEmpty: true }, (cell: Cell, _: number) => {
            const columnLength: number = cell.value ? cell.value.toString().length : 0;
            minLength = columnLength > minLength ? columnLength : minLength;
        });

        // Change the given column's width, plus a buffer.
        // Use the column pre-configured width, if it exists.
        column.width = (columns[i].width || minLength) + 4;
    });
};

// Function for creating the directory path (if applicable) and exporing the Workbook to an output file.
export function exportWorkbook(workbook: Workbook, path: string, fileName: string): void {
    if (!FileStream.existsSync(path)) {
        FileStream.mkdirSync(path, { recursive: true });
    }

    (async () => await workbook.xlsx.writeFile(`${path}/${fileName}.xlsx`))();
};
