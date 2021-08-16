import Columns from "../types/Columns";

const COLUMNS_INVOICE = [
    Columns.Account, Columns.Batch_Number, Columns.Account_Number,
    Columns.Weight_Charge, Columns.DT, Columns.Returns_Other, Columns.Total,
    Columns.Paid_Credit_Card, Columns.Invoice_Due, Columns.Invoice_Sent,
    Columns.AWB, Columns.Pickup_Dates
];

const COLUMNS_PARTNER_STONE_EDGE_SUMMARY = [
    Columns.Parent_Customer_Name, Columns.Customer_Name,
    Columns.USPS_Priority, Columns.USPS_First_Class, Columns.International, Columns.Refund, Columns.APV,
    Columns.CP, Columns.Customer_Cost, Columns.Merchant_Fee, Columns.Rebate
];

const COLUMNS_RETURNS = [
    Columns.Batch_Number, Columns.Description, Columns.Amount
];

const COLUMNS_SUMMARY = [
    Columns.Top_Level_Parent, Columns.Parent_Customer_Name, Columns.Customer_Name,
    Columns.USPS_Priority, Columns.USPS_First_Class, Columns.International, Columns.Refund, Columns.APV,
    Columns.CP, Columns.Customer_Cost, Columns.NSA_Cost, Columns.Gross_Margin, Columns.Merchant_Fee,
    Columns.Rebate, Columns.Royalty, Columns.EHUB_Tech_Fee, Columns.RevShare, Columns.Net_Margin,
    Columns.EHUB_Commission, Columns.Postage_Force_Commission, Columns.PF_Ventures_Commission,
    Columns.Hunter_Commission, Columns.Mike_Commission, Columns.Jordan_Commission,
    Columns.Hunter_Comment_Sold_Commission, Columns.Mike_Comment_Sold_Commission,
    Columns.COR_Commission, Columns.EHUB_Commission_From_PF_Ventures, Columns.Other_Commission,
    Columns.EHUB_Remainder, Columns.Postage_Force_Remainder, Columns.PF_Ventures_Remainder, Columns.Other_Remainder
];

export {
    COLUMNS_INVOICE,
    COLUMNS_PARTNER_STONE_EDGE_SUMMARY,
    COLUMNS_RETURNS,
    COLUMNS_SUMMARY
};
