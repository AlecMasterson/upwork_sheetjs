export default interface Batch {
    Account: string;
    "Account Number": string;
    "AWBs": string;
    "Batch Number": string;
    "Computed Invoice Amount": number;
    "D&T": number;
    "Invoice Due"?: Date;
    "Invoice Number": string;
    "Pickup Dates": string;
    Total: number;
    "Weight Charge": number;
};
