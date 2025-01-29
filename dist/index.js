"use strict";
var __rest = (this && this.__rest) || function (s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ExcelDatabase = void 0;
const XLSX = require("xlsx");
class ExcelDatabase {
    /**
     * Construct Database
     * @param {string} filePath - The file path of the Excel file (e.g., './example.xlsx').
     * @param {string} [sheetName='Sheet1'] - The name of the sheet to use. Defaults to 'Sheet1'.
     */
    constructor(filePath, sheetName = 'Sheet1') {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.data = this.loadData();
    }
    /**
     * Load data from the Excel sheet.
     * @returns {Row[]} The loaded data as an array of objects.
     */
    loadData() {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        return XLSX.utils.sheet_to_json(worksheet);
    }
    /**
     * Save the current data to the Excel sheet.
     */
    saveData() {
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
    select(query = {}) {
        const result = this.data.filter(row => Object.keys(query).every(key => row[key] === query[key]));
        return result.length > 0 ? result : null;
    }
    /**
     * Get a specific column value from a row matching the search criteria.
     * @param {string} searchColumn - The column to search in.
     * @param {any} searchValue - The value to search for.
     * @param {string} targetColumn - The column from which to retrieve the value.
     * @returns {any | undefined} The value from the target column or undefined if not found.
     */
    getColumnValue(searchColumn, searchValue, targetColumn) {
        const row = this.data.find(row => row[searchColumn] === searchValue);
        return row ? row[targetColumn] : undefined;
    }
    /**
     * Insert a new row into the database.
     * @param {Row} newRow - The row to be inserted.
     */
    insert(newRow) {
        this.data.push(newRow);
        this.saveData();
    }
    /**
     * Update existing rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     * @param {Partial<Row>} updateData - The data to update in matched rows.
     */
    update(query, updateData) {
        this.data = this.data.map(row => {
            if (Object.keys(query).every(key => row[key] === query[key])) {
                return Object.assign(Object.assign({}, row), updateData);
            }
            return row;
        });
        this.saveData();
    }
    /**
     * Delete rows that match the query.
     * @param {Partial<Row>} query - The query to match rows.
     */
    delete(query) {
        this.data = this.data.filter(row => !Object.keys(query).every(key => row[key] === query[key]));
        this.saveData();
    }
    /**
     * Add a new sheet to the Excel file.
     * @param {string} sheetName - The name of the new sheet.
     * @param {Row[]} [initialData=[]] - The initial data for the sheet.
     */
    addSheet(sheetName, initialData = []) {
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
    isSheetExists(sheetName) {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames.includes(sheetName) ? 1 : null;
    }
    /**
     * Get all sheet names in the Excel file.
     * @returns {string[]} The list of sheet names.
     */
    getAllSheetNames() {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames;
    }
    /**
     * Get the number of non-empty values in a column.
     * @param {string} columnName - The column name.
     * @returns {number} The count of non-empty values.
     */
    getColumnDatasNumber(columnName) {
        return this.data.filter(row => row[columnName] !== undefined && row[columnName] !== null && row[columnName] !== '').length | 0;
    }
    /**
     * Add a new column to the sheet with a default value.
     * @param {string} columnName - The name of the new column.
     * @param {any} [defaultValue=null] - The default value for the column.
     */
    addColumn(columnName, defaultValue = null) {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        if (!worksheet) {
            throw new Error(`Sheet "${this.sheetName}" does not exist.`);
        }
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
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
    removeColumn(columnName) {
        this.data = this.data.map(row => {
            const _a = row, _b = columnName, _ = _a[_b], remaining = __rest(_a, [typeof _b === "symbol" ? _b : _b + ""]);
            return remaining;
        });
        this.saveData();
    }
}
exports.ExcelDatabase = ExcelDatabase;
