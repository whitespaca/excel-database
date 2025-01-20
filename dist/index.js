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
    constructor(filePath, sheetName = 'Sheet1') {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.data = this.loadData();
    }
    loadData() {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = workbook.Sheets[this.sheetName];
        return XLSX.utils.sheet_to_json(worksheet);
    }
    saveData() {
        const workbook = XLSX.readFile(this.filePath);
        const worksheet = XLSX.utils.json_to_sheet(this.data);
        workbook.Sheets[this.sheetName] = worksheet;
        XLSX.writeFile(workbook, this.filePath);
    }
    select(query = {}) {
        const result = this.data.filter(row => Object.keys(query).every(key => row[key] === query[key]));
        return result.length > 0 ? result : null;
    }
    getColumnValue(searchColumn, searchValue, targetColumn) {
        const row = this.data.find(row => row[searchColumn] === searchValue);
        return row ? row[targetColumn] : undefined;
    }
    insert(newRow) {
        this.data.push(newRow);
        this.saveData();
    }
    update(query, updateData) {
        this.data = this.data.map(row => {
            if (Object.keys(query).every(key => row[key] === query[key])) {
                return Object.assign(Object.assign({}, row), updateData);
            }
            return row;
        });
        this.saveData();
    }
    delete(query) {
        this.data = this.data.filter(row => !Object.keys(query).every(key => row[key] === query[key]));
        this.saveData();
    }
    addSheet(sheetName, initialData = []) {
        const workbook = XLSX.readFile(this.filePath);
        if (workbook.Sheets[sheetName]) {
            throw new Error(`Sheet with name "${sheetName}" already exists.`);
        }
        const worksheet = XLSX.utils.json_to_sheet(initialData);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        XLSX.writeFile(workbook, this.filePath);
    }
    isSheetExists(sheetName) {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames.includes(sheetName) ? 1 : null;
    }
    getAllSheetNames() {
        const workbook = XLSX.readFile(this.filePath);
        return workbook.SheetNames;
    }
    getColumnDatasNumber(columnName) {
        return this.data.filter(row => row[columnName] !== undefined && row[columnName] !== null && row[columnName] !== '').length | 0;
    }
    addColumn(columnName, defaultValue = null) {
        // 1. 워크북 읽기
        const workbook = XLSX.readFile(this.filePath);
        // 2. 현재 시트 가져오기
        const worksheet = workbook.Sheets[this.sheetName];
        if (!worksheet) {
            throw new Error(`Sheet "${this.sheetName}" does not exist.`);
        }
        // 3. 시트 데이터를 JSON 형식으로 변환
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        // 4. 새 열 추가
        if (jsonData.length === 0) {
            // 데이터가 없으면 빈 행을 추가하고 열을 생성
            jsonData.push({ [columnName]: defaultValue });
        }
        else {
            // 데이터가 있으면 각 행에 새 열 추가
            jsonData.forEach(row => {
                if (!(columnName in row)) {
                    row[columnName] = defaultValue;
                }
            });
        }
        // 5. 열 순서 정렬: 기존 열 + 새 열 순서 유지
        const orderedData = jsonData.map(row => {
            const orderedRow = {};
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
    removeColumn(columnName) {
        this.data = this.data.map(row => {
            const _a = row, _b = columnName, _ = _a[_b], remaining = __rest(_a, [typeof _b === "symbol" ? _b : _b + ""]);
            return remaining;
        });
        this.saveData();
    }
}
exports.ExcelDatabase = ExcelDatabase;
