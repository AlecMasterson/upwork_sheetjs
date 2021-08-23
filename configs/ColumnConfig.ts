import ExcelColumns from "../types/ExcelColumns";

const COLUMNS_INVOICE = [
    ExcelColumns.Account, ExcelColumns.Batch_Number, ExcelColumns.Account_Number,
    ExcelColumns.Weight_Charge, ExcelColumns.DT, ExcelColumns.Returns_Other, ExcelColumns.Total,
    ExcelColumns.Paid_Credit_Card, ExcelColumns.Invoice_Due, ExcelColumns.Invoice_Sent,
    ExcelColumns.AWB, ExcelColumns.Pickup_Dates
];

const COLUMNS_PARTNER_STONE_EDGE_SUMMARY = [
    ExcelColumns.Parent_Customer_Name, ExcelColumns.Customer_Name,
    ExcelColumns.USPS_Priority, ExcelColumns.USPS_First_Class, ExcelColumns.International, ExcelColumns.Refund, ExcelColumns.APV,
    ExcelColumns.CP, ExcelColumns.Customer_Cost, ExcelColumns.Merchant_Fee, ExcelColumns.Rebate
];

const COLUMNS_RETURNS = [
    ExcelColumns.Batch_Number, ExcelColumns.Description, ExcelColumns.Amount
];

const COLUMNS_SUMMARY = [
    ExcelColumns.Top_Level_Parent, ExcelColumns.Parent_Customer_Name, ExcelColumns.Customer_Name,
    ExcelColumns.USPS_Priority, ExcelColumns.USPS_First_Class, ExcelColumns.International, ExcelColumns.Refund, ExcelColumns.APV,
    ExcelColumns.CP, ExcelColumns.Customer_Cost, ExcelColumns.NSA_Cost, ExcelColumns.Gross_Margin, ExcelColumns.Merchant_Fee,
    ExcelColumns.Rebate, ExcelColumns.Royalty, ExcelColumns.EHUB_Tech_Fee, ExcelColumns.RevShare, ExcelColumns.Net_Margin,
    ExcelColumns.EHUB_Commission, ExcelColumns.Postage_Force_Commission, ExcelColumns.PF_Ventures_Commission,
    ExcelColumns.Hunter_Commission, ExcelColumns.Mike_Commission, ExcelColumns.Jordan_Commission,
    ExcelColumns.Hunter_Comment_Sold_Commission, ExcelColumns.Mike_Comment_Sold_Commission,
    ExcelColumns.COR_Commission, ExcelColumns.EHUB_Commission_From_PF_Ventures, ExcelColumns.Other_Commission,
    ExcelColumns.EHUB_Remainder, ExcelColumns.Postage_Force_Remainder, ExcelColumns.PF_Ventures_Remainder, ExcelColumns.Other_Remainder
];

export {
    COLUMNS_INVOICE,
    COLUMNS_PARTNER_STONE_EDGE_SUMMARY,
    COLUMNS_RETURNS,
    COLUMNS_SUMMARY
};
