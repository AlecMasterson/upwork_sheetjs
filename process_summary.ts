import ExcelJS, { Workbook, Worksheet } from "exceljs";
import { COLUMNS_PARTNER_STONE_EDGE_SUMMARY, COLUMNS_SUMMARY } from "./configs/ColumnConfig";
import { ExcelColumn } from "./types/ExcelColumns";
import RawData from "./types/RawSummary";
import { add_formatted_table, exportWorkbook } from "./utils/Util";
import * as FileStream from "fs";
import CsvParser from "csv-parser";
import Lodash from "lodash";
import moment from "moment";

/* =================================================
====================================================
CUSTOM TYPES
====================================================
================================================= */

// Custom type to represent each row in the final output file.
type Summary = { [key: string]: number | string };

// Custom type to define the "key -> Summary" relationship.
// This is used to organize the data as it is being parsed.
type SummaryMap = { [key: string]: Summary };

// Custom type to configure each Worksheet in the final output file.
type SheetConfig = { columns: ExcelColumn[], getRows: () => SummaryMap, name: string };

/* =================================================
====================================================
CONFIGURATION AND PREPARATION
====================================================
================================================= */

// The large CSV input file.
const INPUT_FILE_NAME = "data/all_data.csv";

// These contain the summaries of the parsed data for each of the Worksheets.
// The "CustomerSummary" has an extra layer to dynamically (and separately) contain each unique "top_level_parent".
const TopLevelSummary: SummaryMap = {};
const ParentCustomerSummary: SummaryMap = {};
const CustomerSummary: { [topLevelParent: string]: SummaryMap } = {};
const StoneEdgeSummary: SummaryMap = {};

// Each Worksheet needs to be configured here.
const SheetConfigs: SheetConfig[] = [
    { columns: COLUMNS_SUMMARY, getRows: () => TopLevelSummary, name: "Top Level Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => ParentCustomerSummary, name: "Parent Customer Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => CustomerSummary["Ehub"], name: "EHUB Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => CustomerSummary["Postage Force"], name: "Postage Force Summary" },
    { columns: COLUMNS_SUMMARY, getRows: () => CustomerSummary["PF Ventures"], name: "PF Ventures Summary" },
    { columns: COLUMNS_PARTNER_STONE_EDGE_SUMMARY, getRows: () => StoneEdgeSummary, name: "Stone Edge" },

];

/* =================================================
====================================================
FUNCTIONS
====================================================
================================================= */

// Function used to magically update the Summary object based on the column formula configurations.
function update_calculations(summary: Summary, row: RawData, columns = COLUMNS_SUMMARY): Summary {
    columns.filter((column: ExcelColumn) => column.f).forEach((column: ExcelColumn) => {
        if (column.f instanceof Object && Object.keys(column.f).every((key: string) => {
            return column.f[key].startsWith("~") ? row[key] !== column.f[key].substring(1) : row[key] === column.f[key];
        })) {
            (summary[column.key] as number) += 1;
        } else if (column.f === "SUM") {
            (summary[column.key] as number) += parseFloat(row[column.key] as unknown as string);
        }
    });

    return summary;
};

// Function used to individually parse each row from the large input CSV file.
// By parsing each row "one at a time", we prevent overloading the computer's memory and crashing.
function parse_row(row: RawData): void {
    const topLevelParent: string = row[Object.keys(row)[0]] as unknown as string;
    const parentCustomerName: string = row.parent_customer_name || "";
    let customerName: string = row.customer_name || "";

    // Create an empty Summary object with zeroes for all columns.
    const create_empty_summary = (columns: ExcelColumn[]): Summary => {
        const summary: Summary = columns.reduce((r: Summary, column: ExcelColumn) => ({ ...r, [column.key]: 0 }), {});
        return { ...summary, top_level_parent: topLevelParent, parent_customer_name: parentCustomerName, customer_name: customerName };
    };

    // Update the TopLevelSummary with the current row's data. Or create a new Summary object if it doesn't exist.
    // The new Summary object (if applicable) will also be updated with the current row's data.
    Lodash.update(TopLevelSummary, topLevelParent, (summary: Summary) => {
        const newSummary: Summary = update_calculations(summary || create_empty_summary(COLUMNS_SUMMARY), row);
        return { ...newSummary, parent_customer_name: "", customer_name: "" };
    });

    // Update the ParentCustomerSummary with the current row's data. Or create a new Summary object if it doesn't exist.
    // The new Summary object (if applicable) will also be updated with the current row's data.
    // Don't add any row that doesn't have a "parent_customer_name".
    if (parentCustomerName !== "") {
        Lodash.update(ParentCustomerSummary, `${topLevelParent}-${parentCustomerName}`, (summary: Summary) => {
            const newSummary: Summary = update_calculations(summary || create_empty_summary(COLUMNS_SUMMARY), row);
            return { ...newSummary, customer_name: "" };
        });
    }

    // Update the StoneEdgeSummary with the current row's data. Or create a new Summary object if it doesn't exist.
    // The new Summary object (if applicable) will also be updated with the current row's data.
    // Only add a row that has a "parent_customer_name" equal to "Stone Edge".
    if (parentCustomerName === "Stone Edge") {
        Lodash.update(StoneEdgeSummary, `${parentCustomerName}-${customerName}`, (summary: Summary) => {
            return update_calculations(summary || create_empty_summary(COLUMNS_SUMMARY), row);
        });
    }

    // Update the CustomerSummary with the current row's data. Or create a new Summary object if it doesn't exist.
    // The new Summary object (if applicable) will also be updated with the current row's data.
    // Only keep the customer name if the "parent_customer_name" is empty.
    customerName = parentCustomerName === "" ? customerName : "";
    Lodash.update(CustomerSummary, `${topLevelParent}.${parentCustomerName}-${customerName}`, (summary: Summary) => {
        return update_calculations(summary || create_empty_summary(COLUMNS_SUMMARY), row);
    });
}

// Function used to export the parsed data to a new XLSX file.
// This function performs all necessary formatting for the final output file.
function export_summary(): void {
    const workbook: Workbook = new ExcelJS.Workbook();

    // For each Worksheet in the configuration, add it to the Workbook and then add/format the data.
    SheetConfigs.forEach((sheetConfig: SheetConfig) => {
        const sheet: Worksheet = workbook.addWorksheet(sheetConfig.name);
        add_formatted_table(sheet, sheetConfig.name.replace(/\s/g, "_"), sheetConfig.columns, Object.values(sheetConfig.getRows()));
    });

    // Export the Workbook to an XLSX file.
    exportWorkbook(workbook, "reports", `${moment().format("yyyy_MM_DD")}_summary`);
}

/* =================================================
====================================================
MAIN PROCESS
====================================================
================================================= */

FileStream.createReadStream(INPUT_FILE_NAME).pipe(CsvParser())
    .on("data", (row: RawData) => parse_row(row))
    .on("end", () => export_summary());
