import ExcelJS, { Workbook, Worksheet } from "exceljs";
import { COLUMNS_PARTNER_STONE_EDGE_SUMMARY, COLUMNS_SUMMARY } from "./configs/ColumnConfig";
import ValueMap from "./types/ValueMap";
import { Summary } from "./types/Summary";
import { ExcelColumn } from "./types/Columns";
import SheetConfigType from "./types/SheetConfigType";
import { adjust_column_widths, exportWorkbook, parseExcelColumn } from "./utils/Util";
import * as FileStream from "fs";
import CsvParser from "csv-parser";
import Lodash from "lodash";
import moment from "moment";

const ParsedData: { [topLevelParent: string]: { [parentCustomerName: string]: { [customerName: string]: ValueMap<object>[] } } } = {};
const SheetConfig: SheetConfigType[] = [
    { columns: COLUMNS_SUMMARY, getRows: getRows_TopLevelSummary, name: "Top Level Summary" },
    { columns: COLUMNS_SUMMARY, getRows: getRows_ParentCustomerSummary, name: "Parent Customer Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => getRows_Summary("Ehub"), name: "EHUB Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => getRows_Summary("Postage Force"), name: "Postage Force Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => getRows_Summary("PF Ventures"), name: "PF Ventures Summary" },
    { columns: COLUMNS_PARTNER_STONE_EDGE_SUMMARY, getRows: () => getRows_StoneEdge(), name: "Stone Edge" }
];

FileStream.createReadStream(`data/all_data_${moment().format("yyyy_MM_DD")}.csv`).pipe(CsvParser())
    .on("data", (row: ValueMap<object>) => parse_row(row))
    .on("end", () => export_summary());

function parse_row(row: ValueMap<object>): void {
    const topLevelParent: string = row[Object.keys(row)[0]] as unknown as string;
    const parentCustomerName: string = (row["parent_customer_name"] || "") as unknown as string;
    const customerName: string = (row["customer_name"] || "") as unknown as string;

    if (!Lodash.has(ParsedData, [topLevelParent, parentCustomerName, customerName])) {
        Lodash.set(ParsedData, [topLevelParent, parentCustomerName, customerName], []);
    }

    ParsedData[topLevelParent][parentCustomerName][customerName].push(row);
}

function export_summary(): void {
    const workbook: Workbook = new ExcelJS.Workbook();

    SheetConfig.forEach((sheetConfig: SheetConfigType) => {
        const sheet: Worksheet = workbook.addWorksheet(sheetConfig.name);

        sheet.addTable({
            columns: sheetConfig.columns.map((column: ExcelColumn) => parseExcelColumn(column)),
            name: sheetConfig.name.replace(/\s/g, "_"),
            ref: "A1",
            rows: sheetConfig.getRows().map((row: Summary) => sheetConfig.columns.map((column: ExcelColumn) => row[column.key])),
            style: { showRowStripes: true }
        });

        // Apply cell formatting to each column, excluding the header row.
        sheetConfig.columns.forEach((column: ExcelColumn, i: number) => sheet.getColumn(i + 1).numFmt = Lodash.get(column, "style.numFmt", null));
        sheet.getRow(1).numFmt = null;

        // Apply column formatting, fit the width to the cell data.
        adjust_column_widths(sheet, sheetConfig.columns);
    });

    // Export the Workbook to an XLSX file.
    exportWorkbook(workbook, "reports", `${moment().format("yyyy_MM_DD")}_summary`);
}

function getRows_TopLevelSummary(): Summary[] {
    return Object.keys(ParsedData).map((topLevelParent: string) => {
        const result: Summary = COLUMNS_SUMMARY.reduce((r: Summary, column: ExcelColumn) => ({ ...r, [column.key]: 0 }), {});

        Object.keys(ParsedData[topLevelParent]).forEach((parentCustomerName: string) => {
            Object.keys(ParsedData[topLevelParent][parentCustomerName]).forEach((customerName: string) => {
                ParsedData[topLevelParent][parentCustomerName][customerName].forEach((row: ValueMap<object>) => update_calculations(result, row));
            });
        });

        return { ...result, top_level_parent: topLevelParent, parent_customer_name: "", customer_name: "" };
    });
}

function getRows_ParentCustomerSummary(): Summary[] {
    return Lodash.flattenDeep(Object.keys(ParsedData).map((topLevelParent: string) => {
        return Object.keys(ParsedData[topLevelParent]).map((parentCustomerName: string) => {
            const result: Summary = COLUMNS_SUMMARY.reduce((r: Summary, column: ExcelColumn) => ({ ...r, [column.key]: 0 }), {});

            Object.keys(ParsedData[topLevelParent][parentCustomerName]).forEach((customerName: string) => {
                ParsedData[topLevelParent][parentCustomerName][customerName].forEach((row: ValueMap<object>) => update_calculations(result, row));
            });

            return { ...result, top_level_parent: topLevelParent, parent_customer_name: parentCustomerName, customer_name: "" };
        });
    })).filter((row: Summary) => row["parent_customer_name"] !== "");
}

function getRows_Summary(topLevelParent: string): Summary[] {
    const parentCustomerSummary: Summary[] = getRows_ParentCustomerSummary().filter((row: Summary) => row["top_level_parent"] === topLevelParent);

    const customerSummary: Summary[] = Lodash.flattenDeep(Object.keys(ParsedData[topLevelParent])
        .filter((parentCustomerName: string) => parentCustomerName === "").map((parentCustomerName: string) => {
            return Object.keys(ParsedData[topLevelParent][parentCustomerName]).map((customerName: string) => {
                const result: Summary = COLUMNS_SUMMARY.reduce((r: Summary, column: ExcelColumn) => ({ ...r, [column.key]: 0 }), {});

                ParsedData[topLevelParent][parentCustomerName][customerName].forEach((row: ValueMap<object>) => update_calculations(result, row));

                return { ...result, top_level_parent: topLevelParent, parent_customer_name: parentCustomerName, customer_name: customerName };
            });
        }));

    return Lodash.flatten([parentCustomerSummary, customerSummary]);
}

function getRows_StoneEdge(): Summary[] {
    return Lodash.flattenDeep(Object.keys(ParsedData).map((topLevelParent: string) => {
        return Object.keys(ParsedData[topLevelParent]["Stone Edge"] || []).filter((customerName: string) => customerName !== "").map((customerName: string) => {
            const result: Summary = COLUMNS_PARTNER_STONE_EDGE_SUMMARY.reduce((r: Summary, column: ExcelColumn) => ({ ...r, [column.key]: 0 }), {});

            ParsedData[topLevelParent]["Stone Edge"][customerName].forEach((row: ValueMap<object>) => update_calculations(result, row, COLUMNS_PARTNER_STONE_EDGE_SUMMARY));

            return { ...result, parent_customer_name: "Stone Edge", customer_name: customerName };
        });
    }));
}

function update_calculations(result: Summary, row: ValueMap<object>, columns = COLUMNS_SUMMARY): void {
    columns.filter((column: ExcelColumn) => column.f).forEach((column: ExcelColumn) => {
        if (column.f instanceof Object && Object.keys(column.f).every((key: string) => {
            return column.f[key].startsWith("~") ? row[key] !== column.f[key].substring(1) : row[key] === column.f[key];
        })) {
            (result[column.key] as number) += 1;
        } else if (column.f === "SUM") {
            (result[column.key] as number) += parseFloat(row[column.key] as unknown as string);
        }
    });
}
