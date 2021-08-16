import { Style } from "exceljs";
import { FORMAT_CURRENCY, FORMAT_NUMBER, MIN_COLUMN_WIDTH_DATE } from "../configs/StyleConfig";

export type TotalsRowFunctionType = "none" | "average" | "countNums" | "count" | "max" | "min" | "stdDev" | "var" | "sum" | "custom";

export interface ExcelColumn {
    f?: (object | string);
    filterButton?: boolean;
    key: string;
    name: string;
    style?: Partial<Style>;
    totalsRowFunction?: TotalsRowFunctionType;
    totalsRowLabel?: string;
    width?: number;
};

const ExcelColumns: { [key: string]: ExcelColumn } = {
    Account: { key: "Account", name: "Account", totalsRowLabel: "Total" },
    Account_Number: { key: "Account Number", name: "Account Number" },
    Amount: { key: "Amount", name: "Amount", style: { numFmt: FORMAT_CURRENCY } },
    APV: { f: { status: "APV" }, key: "APV", name: "APV", style: { numFmt: FORMAT_NUMBER } },
    AWB: { key: "AWBs", name: "AWBs", },
    Batch_Number: { key: "Batch Number", name: "Batch Number" },
    COR_Commission: { f: "SUM", key: "cor_commission", name: "COR Commission", style: { numFmt: FORMAT_CURRENCY } },
    CP: { f: "SUM", key: "cp", name: "CP", style: { numFmt: FORMAT_CURRENCY } },
    Customer_Cost: { f: "SUM", key: "customer_cost", name: "Customer Cost", style: { numFmt: FORMAT_CURRENCY } },
    Customer_Name: { filterButton: true, key: "customer_name", name: "Customer Name" },
    Description: { key: "Description", name: "Description" },
    DT: { key: "D&T", name: "D&T", style: { numFmt: FORMAT_CURRENCY } },
    EHUB_Commission: { f: "SUM", key: "ehub_commission", name: "EHUB Commission", style: { numFmt: FORMAT_CURRENCY } },
    EHUB_Commission_From_PF_Ventures: { f: "SUM", key: "ehub_commission_from_pfventures", name: "EHUB Commission from PF Ventures", style: { numFmt: FORMAT_CURRENCY } },
    EHUB_Remainder: { f: "SUM", key: "ehub_remainder", name: "EHUB Remainder", style: { numFmt: FORMAT_CURRENCY } },
    EHUB_Tech_Fee: { f: "SUM", key: "ehub_tech_fee", name: "EHUB Tech Fee", style: { numFmt: FORMAT_CURRENCY } },
    Gross_Margin: { f: "SUM", key: "gross_margin", name: "Gross Margin", style: { numFmt: FORMAT_CURRENCY } },
    Hunter_Comment_Sold_Commission: { f: "SUM", key: "hunter_comment_sold_commission", name: "Hunter Comment Sold Commission", style: { numFmt: FORMAT_CURRENCY } },
    Hunter_Commission: { f: "SUM", key: "hunter_commission", name: "Hunter Commission", style: { numFmt: FORMAT_CURRENCY } },
    International: { f: { to_country: "~US", status: "shipped" }, key: "International", name: "International", style: { numFmt: FORMAT_NUMBER } },
    Invoice_Due: { key: "Invoice Due", name: "Invoice Due", style: { numFmt: "dd-mmm" }, width: MIN_COLUMN_WIDTH_DATE },
    Invoice_Sent: { key: "Invoice Sent", name: "Invoice Sent", style: { numFmt: "dd-mmm" }, width: MIN_COLUMN_WIDTH_DATE },
    Jordan_Commission: { f: "SUM", key: "jordan_commission", name: "Jordan Commission", style: { numFmt: FORMAT_CURRENCY } },
    Merchant_Fee: { f: "SUM", key: "merchant_fee", name: "Merchant Fee", style: { numFmt: FORMAT_CURRENCY } },
    Mike_Comment_Sold_Commission: { f: "SUM", key: "mike_comment_sold_commission", name: "Mike Comment Sold Commission", style: { numFmt: FORMAT_CURRENCY } },
    Mike_Commission: { f: "SUM", key: "mike_commission", name: "Mike Commission", style: { numFmt: FORMAT_CURRENCY } },
    Net_Margin: { f: "SUM", key: "net_margin", name: "Net Margin", style: { numFmt: FORMAT_CURRENCY } },
    NSA_Cost: { f: "SUM", key: "nsa_cost", name: "NSA Cost", style: { numFmt: FORMAT_CURRENCY } },
    Other_Commission: { f: "SUM", key: "other_commission", name: "Other Commission", style: { numFmt: FORMAT_CURRENCY } },
    Other_Remainder: { f: "SUM", key: "other_remainder", name: "Other Remainder", style: { numFmt: FORMAT_CURRENCY } },
    Paid_Credit_Card: { key: "PaidCreditCard", name: "If Paid By Credit Card", style: { numFmt: FORMAT_CURRENCY }, totalsRowFunction: "sum" },
    Parent_Customer_Name: { filterButton: true, key: "parent_customer_name", name: "Parent Customer Name" },
    PF_Ventures_Commission: { f: "SUM", key: "pfventures_commission", name: "PF Ventures Commission", style: { numFmt: FORMAT_CURRENCY } },
    PF_Ventures_Remainder: { f: "SUM", key: "pfventures_remainder", name: "PF Ventures Remainder", style: { numFmt: FORMAT_CURRENCY } },
    Pickup_Dates: { key: "Pickup Dates", name: "Pickup Dates", },
    Postage_Force_Commission: { f: "SUM", key: "postage_force_commission", name: "Postage Force Commission", style: { numFmt: FORMAT_CURRENCY } },
    Postage_Force_Remainder: { f: "SUM", key: "postage_force_remainder", name: "Postage Force Remainder", style: { numFmt: FORMAT_CURRENCY } },
    Rebate: { f: "SUM", key: "rebate", name: "Rebate", style: { numFmt: FORMAT_CURRENCY } },
    Refund: { f: { status: "cancelled" }, key: "refund", name: "Refund", style: { numFmt: FORMAT_NUMBER } },
    Returns_Other: { key: "Returns&Other", name: "Returns & Other", style: { numFmt: FORMAT_CURRENCY } },
    RevShare: { f: "SUM", key: "revshare", name: "Revshare", style: { numFmt: FORMAT_CURRENCY } },
    Royalty: { f: "SUM", key: "royalty", name: "Royalty", style: { numFmt: FORMAT_CURRENCY } },
    Top_Level_Parent: { filterButton: true, key: "top_level_parent", name: "Top Level Parent" },
    Total: { key: "Total", name: "Total", style: { numFmt: FORMAT_CURRENCY }, totalsRowFunction: "sum" },
    USPS_First_Class: { f: { shipment_type: "USPS First Class", status: "shipped" }, key: "USPS First Class", name: "USPS First Class", style: { numFmt: FORMAT_NUMBER } },
    USPS_Priority: { f: { shipment_type: "USPS Priority", status: "shipped" }, key: "USPS Priority", name: "USPS Priority", style: { numFmt: FORMAT_NUMBER } },
    Weight_Charge: { key: "Weight Charge", name: "Weight Charge", style: { numFmt: FORMAT_CURRENCY } },
};

export default ExcelColumns;
