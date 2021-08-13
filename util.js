const Config = require('./config.js');

function add_formulas(sheet, data) {
    // formulas for the column "Returns & Other".
    sheet.getColumn('F').eachCell((cell, index) => {
        if (index === 1) return;
        cell.value = { formula: `SUMIFS(returns_and_other!C:C, returns_and_other!A:A, B${index})` };
    })

    // Add formulas for the columns "Total" and "If Paid By Credit Card".
    sheet.fillFormula(`G2:G${data.length + 1}`, 'D2+E2+F2');
    sheet.fillFormula(`H2:H${data.length + 1}`, 'G2*1.03');

    // Add formula for the column "Invoice Due".
    sheet.fillFormula(`J2:J${data.length + 1}`, 'I2+7');

    // Add formulas for the total row on the "Total" and "If Paid By Credit Card" columns.
    sheet.getCell(`G${data.length + 2}`).value = { formula: `SUM(G2:G${data.length + 1})` };
    sheet.getCell(`H${data.length + 2}`).value = { formula: `SUM(H2:H${data.length + 1})` };

    // Add bolded formatting to the total row.
    sheet.getRow(data.length + 2).font = { bold: true };
}

function adjust_column_widths(sheet, column_config) {
    sheet.columns.filter((column, i) => !column_config[i].width).forEach((column) => {
        let maxLength = Config.MIN_COLUMN_WIDTH;

        column.eachCell({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? cell.value.toString().length : 0;
            maxLength = columnLength > maxLength ? columnLength : maxLength;
        });

        column.width = maxLength;
    });
}

function createSheets(workbook) {
    const sheetInvoice = workbook.addWorksheet('invoice');
    const sheetReturnsAndOther = workbook.addWorksheet('returns_and_other');

    // Create the Google Sheets based on the configuration.
    sheetInvoice.columns = Config.COLUMNS_INVOICE;
    sheetReturnsAndOther.columns = Config.COLUMNS_RETURNS;

    // Add bolded formatting to the header row.
    sheetInvoice.getRow(1).font = { bold: true };
    sheetReturnsAndOther.getRow(1).font = { bold: true };

    return { invoice: sheetInvoice, returns: sheetReturnsAndOther };
}

module.exports = {
    add_formulas,
    adjust_column_widths,
    createSheets
};
