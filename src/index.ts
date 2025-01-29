import * as XLSX from 'xlsx';

interface Row {
    [key: string]: any;
}

export class ExcelDatabase {
    private filePath: string;
    private sheetName: string;
    private data: Row[];

    /**
     * Construct Database
     * @param {string} filePath - The file path of the Excel file (e.g., './example.xlsx').
     * @param {string} [sheetName='Sheet1'] - The name of the sheet to use. Defaults to 'Sheet1'.
     */
    constructor(filePath: string, sheetName: string = 'Sheet1') {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.data = this.loadData();
    }

    /**
     * Load data from the Excel sheet.
     * @returns {Row[]} The loaded data as an array of objects.
     */
    private loadData(): Row[] {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        return XLSX.utils.sheet_to_json<Row>(worksheet);
    }

    /**
     * Save the current data to the Excel sheet.
     */
    private saveData(): void {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = XLSX.utils.json_to_sheet(this.data);
        workbook.Sheets[this.sheetName] = worksheet;
        XLSX.writeFile(workbook, this.filePath);
    }

    /**
     * Select values based on a query.
     * @param {Partial<Row>} query - The query object containing key-value pairs to match.
     * @returns {Row[] | null} The matched rows or null if no match is found.
     */
    public select(query: Partial<Row> = {}): Row[] | null {
        const result = this.data.filter(row =>
            Object.keys(query).every(key => row[key] === query[key])
        );
        return result.length > 0 ? result : null;
    }

    /**
     * Get a specific column value from a row matching the search criteria.
     * @param {string} searchColumn - The column to search in.
     * @param {any} searchValue - The value to search for.
     * @param {string} targetColumn - The column from which to retrieve the value.
     * @returns {any | undefined} The value from the target column or undefined if not found.
     */
    public getColumnValue(searchColumn: string, searchValue: any, targetColumn: string): any | undefined {
        const row = this.data.find(row => row[searchColumn] === searchValue);
        return row ? row[targetColumn] : undefined;
    }

    /**
     * Insert a new row into the database.
     * @param {Row} newRow - The row to be inserted.
     */
    public insert(newRow: Row): void {
        this.data.push(newRow);
        this.saveData();
    }

    /**
     * Update existing rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     * @param {Partial<Row>} updateData - The data to update in matched rows.
     */
    public update(query: Partial<Row>, updateData: Partial<Row>): void {
        this.data = this.data.map(row => {
            if (Object.keys(query).every(key => row[key] === query[key])) {
                return { ...row, ...updateData };
            }
            return row;
        });
        this.saveData();
    }

    /**
     * Delete rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     */
    public delete(query: Partial<Row>): void {
        this.data = this.data.filter(row =>
            !Object.keys(query).every(key => row[key] === query[key])
        );
        this.saveData();
    }

    /**
     * Add a new sheet to the Excel file.
     * @param {string} sheetName - The name of the new sheet.
     * @param {Row[]} [initialData=[]] - The initial data for the sheet.
     */
    public addSheet(sheetName: string, initialData: Row[] = []): void {
        const workbook = XLSX.readFile(this.filePath);
        if (workbook.Sheets[sheetName]) {
            throw new Error(`Sheet with name "${sheetName}" already exists.`);
        }
        const worksheet = XLSX.utils.json_to_sheet(initialData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        XLSX.writeFile(workbook, this.filePath);
    }

    /**
     * Check if a sheet exists.
     * @param {string} sheetName - The name of the sheet to check.
     * @returns {1 | null} 1 if the sheet exists, null otherwise.
     */
    public isSheetExists(sheetName: string): 1 | null {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames.includes(sheetName) ? 1 : null;
    }

    /**
     * Get all sheet names in the Excel file.
     * @returns {string[]} The list of sheet names.
     */
    public getAllSheetNames(): string[] {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames;
    }

    /**
     * Get the number of non-empty values in a column.
     * @param {string} columnName - The column name.
     * @returns {number} The count of non-empty values.
     */
    public getColumnDatasNumber(columnName: string): number {
        return this.data.filter(row => row[columnName] !== undefined && row[columnName] !== null && row[columnName] !== '').length | 0;
    }

    /**
     * Add a new column to the sheet with a default value.
     * @param {string} columnName - The name of the new column.
     * @param {any} [defaultValue=null] - The default value for the column.
     */
    public addColumn(columnName: string, defaultValue: any = null): void {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        if (!worksheet) {
            throw new Error(`Sheet "${this.sheetName}" does not exist.`);
        }
        const jsonData: Row[] = XLSX.utils.sheet_to_json<Row>(worksheet);
        jsonData.forEach(row => {
            if (!(columnName in row)) {
                row[columnName] = defaultValue;
            }
        });
        const updatedWorksheet = XLSX.utils.json_to_sheet(jsonData);
        workbook.Sheets[this.sheetName] = updatedWorksheet;
        XLSX.writeFile(workbook, this.filePath);
    }

    /**
     * Remove a column from all rows.
     * @param {string} columnName - The column to remove.
     */
    public removeColumn(columnName: string): void {
        this.data = this.data.map(row => {
            const { [columnName]: _, ...remaining } = row;
            return remaining;
        });
        this.saveData();
    }
}
