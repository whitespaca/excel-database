import * as XLSX from 'xlsx';

interface Row {
    [key: string]: any;
}

export class ExcelDatabase {
    private filePath: string;
    private sheetName: string; // In SQL: Table
    private data: Row[];

    constructor(filePath: string, sheetName: string = 'Sheet1') {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.data = this.loadData();
    }

    private loadData(): Row[] {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        return XLSX.utils.sheet_to_json<Row>(worksheet);
    }

    private saveData() {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = XLSX.utils.json_to_sheet(this.data);
        workbook.Sheets[this.sheetName] = worksheet;
        XLSX.writeFile(workbook, this.filePath);
    }

    public select(query: Partial<Row> = {}): Row[] | null {
        const result = this.data.filter(row =>
            Object.keys(query).every(key => row[key] === query[key])
        );
        return result.length > 0 ? result : null;
    }

    public getColumnValue(searchColumn: string, searchValue: any, targetColumn: string): any | undefined {
        const row = this.data.find(row => row[searchColumn] === searchValue);
        return row ? row[targetColumn] : undefined;
    }

    public insert(newRow: Row) {
        this.data.push(newRow);
        this.saveData();
    }

    public update(query: Partial<Row>, updateData: Partial<Row>) {
        this.data = this.data.map(row => {
            if (Object.keys(query).every(key => row[key] === query[key])) {
                return { ...row, ...updateData };
            }
            return row;
        });
        this.saveData();
    }

    public delete(query: Partial<Row>) {
        this.data = this.data.filter(row =>
            !Object.keys(query).every(key => row[key] === query[key])
        );
        this.saveData();
    }

    public addSheet(sheetName: string, initialData: Row[] = []) {
        const workbook = XLSX.readFile(this.filePath);
        if (workbook.Sheets[sheetName]) {
            throw new Error(`Sheet with name "${sheetName}" already exists.`);
        }
    
        const worksheet = XLSX.utils.json_to_sheet(initialData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        XLSX.writeFile(workbook, this.filePath);
    }

    public isSheetExists(sheetName: string): 1 | null {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames.includes(sheetName) ? 1 : null;
    }

    public getAllSheetNames(): string[] {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames;
    }

    public getColumnDatasNumber(columnName: string): number {
        return this.data.filter(row => row[columnName] !== undefined && row[columnName] !== null && row[columnName] !== '').length | 0;
    }

    public addColumn(columnName: string, defaultValue: any = null) {
        const workbook = XLSX.readFile(this.filePath);
    
        const worksheet = workbook.Sheets[this.sheetName];
        if (!worksheet) {
            throw new Error(`Sheet "${this.sheetName}" does not exist.`);
        }
    
        const jsonData: Row[] = XLSX.utils.sheet_to_json<Row>(worksheet);
    
        if (jsonData.length === 0) {
            jsonData.push({ [columnName]: defaultValue });
        } else {
            jsonData.forEach(row => {
                if (!(columnName in row)) {
                    row[columnName] = defaultValue;
                }
            });
        }
    
        const orderedData = jsonData.map(row => {
            const orderedRow: any = {};
            const columns = Object.keys(row).filter(key => key !== columnName);
            columns.forEach(key => {
                orderedRow[key] = row[key];
            });
            orderedRow[columnName] = row[columnName];
            return orderedRow;
        });
    
        const updatedWorksheet = XLSX.utils.json_to_sheet(orderedData);
        workbook.Sheets[this.sheetName] = updatedWorksheet;
    
        XLSX.writeFile(workbook, this.filePath);
    }
    
    
    public removeColumn(columnName: string) {
        this.data = this.data.map(row => {
            const { [columnName]: _, ...remaining } = row;
            return remaining;
        });
        this.saveData();
    }
}