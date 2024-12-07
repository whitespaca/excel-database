import * as XLSX from 'xlsx';

interface Row {
    [key: string]: any;
}

export class ExcelDatabase {
    private filePath: string;
    private sheetName: string;
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
        const worksheet = XLSX.utils.json_to_sheet(this.data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, this.sheetName);
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
        this.saveData();
    }

    public isSheetExists(sheetName: string): 1 | null {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames.includes(sheetName) ? 1 : null;
    }
}