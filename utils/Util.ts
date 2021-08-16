import { Cell, Column, TableColumnProperties, Workbook, Worksheet } from "exceljs";
import { MIN_COLUMN_WIDTH } from "../configs/StyleConfig";
import { ExcelColumn } from "../types/Columns";
import * as FileStream from "fs";

export function adjust_column_widths(sheet: Worksheet, columns: ExcelColumn[]): void {
    sheet.columns.forEach((column: Partial<Column>, i: number) => {
        let minLength: number = MIN_COLUMN_WIDTH;

        column.eachCell && column.eachCell({ includeEmpty: true }, (cell: Cell, _: number) => {
            const columnLength: number = cell.value ? cell.value.toString().length : 0;
            minLength = columnLength > minLength ? columnLength : minLength;
        });

        column.width = columns[i].width || minLength + 4;
    });
};

export function exportWorkbook(workbook: Workbook, path: string, fileName: string): void {
    // Create the directory if it does not exist.
    if (!FileStream.existsSync(path)) {
        FileStream.mkdirSync(path, { recursive: true });
    }

    (async () => await workbook.xlsx.writeFile(`${path}/${fileName}.xlsx`))();
};

export function parseExcelColumn(column: ExcelColumn): TableColumnProperties {
    return {
        name: column.name,
        filterButton: column.filterButton,
        totalsRowFunction: column.totalsRowFunction,
        totalsRowLabel: column.totalsRowLabel
    };
};
