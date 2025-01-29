interface Row {
    [key: string]: any;
}
export declare class ExcelDatabase {
    private filePath;
    private sheetName;
    private data;
    /**
     * Construct Database
     * @param {string} filePath - The file path of the Excel file (e.g., './example.xlsx').
     * @param {string} [sheetName='Sheet1'] - The name of the sheet to use. Defaults to 'Sheet1'.
     */
    constructor(filePath: string, sheetName?: string);
    /**
     * Load data from the Excel sheet.
     * @returns {Row[]} The loaded data as an array of objects.
     */
    private loadData;
    /**
     * Save the current data to the Excel sheet.
     */
    private saveData;
    /**
     * Select values based on a query.
     * @param {Partial<Row>} query - The query object containing key-value pairs to match.
     * @returns {Row[] | null} The matched rows or null if no match is found.
     */
    select(query?: Partial<Row>): Row[] | null;
    /**
     * Get a specific column value from a row matching the search criteria.
     * @param {string} searchColumn - The column to search in.
     * @param {any} searchValue - The value to search for.
     * @param {string} targetColumn - The column from which to retrieve the value.
     * @returns {any | undefined} The value from the target column or undefined if not found.
     */
    getColumnValue(searchColumn: string, searchValue: any, targetColumn: string): any | undefined;
    /**
     * Insert a new row into the database.
     * @param {Row} newRow - The row to be inserted.
     */
    insert(newRow: Row): void;
    /**
     * Update existing rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     * @param {Partial<Row>} updateData - The data to update in matched rows.
     */
    update(query: Partial<Row>, updateData: Partial<Row>): void;
    /**
     * Delete rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     */
    delete(query: Partial<Row>): void;
    /**
     * Add a new sheet to the Excel file.
     * @param {string} sheetName - The name of the new sheet.
     * @param {Row[]} [initialData=[]] - The initial data for the sheet.
     */
    addSheet(sheetName: string, initialData?: Row[]): void;
    /**
     * Check if a sheet exists.
     * @param {string} sheetName - The name of the sheet to check.
     * @returns {1 | null} 1 if the sheet exists, null otherwise.
     */
    isSheetExists(sheetName: string): 1 | null;
    /**
     * Get all sheet names in the Excel file.
     * @returns {string[]} The list of sheet names.
     */
    getAllSheetNames(): string[];
    /**
     * Get the number of non-empty values in a column.
     * @param {string} columnName - The column name.
     * @returns {number} The count of non-empty values.
     */
    getColumnDatasNumber(columnName: string): number;
    /**
     * Add a new column to the sheet with a default value.
     * @param {string} columnName - The name of the new column.
     * @param {any} [defaultValue=null] - The default value for the column.
     */
    addColumn(columnName: string, defaultValue?: any): void;
    /**
     * Remove a column from all rows.
     * @param {string} columnName - The column to remove.
     */
    removeColumn(columnName: string): void;
}
export {};
