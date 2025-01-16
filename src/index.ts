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
        // 1. 워크북 읽기
        const workbook = XLSX.readFile(this.filePath);
    
        // 2. 현재 시트 가져오기
        const worksheet = workbook.Sheets[this.sheetName];
        if (!worksheet) {
            throw new Error(`Sheet "${this.sheetName}" does not exist.`);
        }
    
        // 3. 시트 데이터를 JSON 형식으로 변환
        const jsonData: Row[] = XLSX.utils.sheet_to_json<Row>(worksheet);
    
        // 4. 새 열 추가
        if (jsonData.length === 0) {
            // 데이터가 없으면 빈 행을 추가하고 열을 생성
            jsonData.push({ [columnName]: defaultValue });
        } else {
            // 데이터가 있으면 각 행에 새 열 추가
            jsonData.forEach(row => {
                if (!(columnName in row)) {
                    row[columnName] = defaultValue;
                }
            });
        }
    
        // 5. 열 순서 정렬: 기존 열 + 새 열 순서 유지
        const orderedData = jsonData.map(row => {
            const orderedRow: any = {};
            const columns = Object.keys(row).filter(key => key !== columnName);
            // 기존 열을 먼저 추가
            columns.forEach(key => {
                orderedRow[key] = row[key];
            });
            // 새 열을 마지막에 추가
            orderedRow[columnName] = row[columnName];
            return orderedRow;
        });
    
        // 6. JSON 데이터를 다시 시트로 변환
        const updatedWorksheet = XLSX.utils.json_to_sheet(orderedData);
        workbook.Sheets[this.sheetName] = updatedWorksheet;
    
        // 7. 워크북 저장 (다른 시트 유지)
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