interface Row {
    [key: string]: any;
}
export declare class ExcelDatabase {
    private filePath;
    private sheetName;
    private data;
    constructor(filePath: string, sheetName?: string);
    private loadData;
    private saveData;
    select(query?: Partial<Row>): Row[] | null;
    getColumnValue(searchColumn: string, searchValue: any, targetColumn: string): any | undefined;
    insert(newRow: Row): void;
    update(query: Partial<Row>, updateData: Partial<Row>): void;
    delete(query: Partial<Row>): void;
    addSheet(sheetName: string, initialData?: Row[]): void;
    isSheetExists(sheetName: string): 1 | null;
    getAllSheetNames(): string[];
    getColumnDatasNumber(columnName: string): number;
    addColumn(columnName: string, defaultValue?: any): void;
    removeColumn(columnName: string): void;
}
export {};
