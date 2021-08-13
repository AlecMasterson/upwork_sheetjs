const CURRENCY_FORMAT = '"$"#,##0.00;[Red]\-"$"#,##0.00';
const MIN_COLUMN_WIDTH = 10;
const MIN_COLUMN_WIDTH_DATES = 15;

const COLUMNS_INVOICE = [
    { header: 'Account', key: 'Account' },
    { header: 'Batch Number', key: 'Batch Number' },
    { header: 'Account Number', key: 'Account Number' },
    { header: 'Weight Charge', key: 'Weight Charge', style: { numFmt: CURRENCY_FORMAT } },
    { header: 'D&T', key: 'D&T', style: { numFmt: CURRENCY_FORMAT } },
    { header: 'Returns & Other', style: { numFmt: CURRENCY_FORMAT } },
    { header: 'Total', style: { numFmt: CURRENCY_FORMAT } },
    { header: 'If Paid By Credit Card', style: { numFmt: CURRENCY_FORMAT } },
    { header: 'Invoice Due', key: 'Invoice Due', style: { numFmt: 'dd-mmm' }, width: MIN_COLUMN_WIDTH_DATES },
    { header: 'Invoice Sent', style: { numFmt: 'dd-mmm' }, width: MIN_COLUMN_WIDTH_DATES },
    { header: 'AWBs', key: 'AWBs' },
    { header: 'Pickup Dates', key: 'Pickup Dates' }
];

const COLUMNS_RETURNS = [
    { header: 'Batch Number' },
    { header: 'Description' },
    { header: 'Amount' }
];

module.exports = {
    COLUMNS_INVOICE: COLUMNS_INVOICE,
    COLUMNS_RETURNS: COLUMNS_RETURNS,
    MIN_COLUMN_WIDTH: MIN_COLUMN_WIDTH
};
