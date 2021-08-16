import { Summary } from "./Summary";
import { ExcelColumn } from "./Columns";

export default interface SheetConfigType {
    columns: ExcelColumn[];
    getRows: () => Summary[];
    name: string;
};
