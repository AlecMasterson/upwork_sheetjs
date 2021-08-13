const fs = require('fs');
const moment = require('moment');
const ExcelJS = require('exceljs');
const Config = require('./config.js');
const Util = require('./util.js');

function processAccount(batches) {
    // Parse the batches in an array of objects containing only the data to be put in the Excel document.
    // Additionally, add the "Invoice Due" data with today's date.
    const rows = batches
        .map(row => ({ ...row, 'Invoice Due': new Date() }))
        .map(row => Config.COLUMNS_INVOICE
            .filter(column => column.key)
            .reduce((r, c) => ({ ...r, [c.key]: row[c.key] }), {}));

    // Create an Excel Workbook with the two sheets inside.
    const workbook = new ExcelJS.Workbook();
    const sheets = Util.createSheets(workbook);

    // Add the data and formulas to the "invoice" sheet.
    sheets.invoice.addRows(rows);
    Util.add_formulas(sheets.invoice, rows);

    // Fit the column widths to the data.
    Util.adjust_column_widths(sheets.invoice, Config.COLUMNS_INVOICE);
    Util.adjust_column_widths(sheets.returns, Config.COLUMNS_RETURNS);

    const date = moment().format('MM-DD-yyyy');
    const dir = `reports/${date}/invoices`;

    // Create the directory if it does not exist.
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }

    // Export the Workbook to an XLSX file.
    (async () => await workbook.xlsx.writeFile(`${dir}/${rows[0]['Account']}_invoice_${date}.xlsx`))();
}

// Read the JSON input data and process each account separately.
const data = JSON.parse(fs.readFileSync('example.json'));
Object.keys(data).forEach(accountNumber => processAccount(data[accountNumber]));
