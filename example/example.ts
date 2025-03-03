import { ExcelDatabase } from 'excel-database';

// file root and sheet name
const db = new ExcelDatabase('../data.xlsx', 'Sheet1');

// SELECT
const selectedRows = db.select({ name: 'John Doe' });
console.log('Selected Rows:', selectedRows);

// INSERT
db.insert({ name: 'Jane Doe', age: 30, city: 'New York' });

// UPDATE
db.update({ name: 'Jane Doe' }, { age: 31 });


// DELETE
db.delete({ name: 'John Doe' });

// getColumnValue
const test = db.getColumnValue('name', 'Jane Doe', 'age');
console.log(test)

// addSheet
const initialData = [
  { name: 'Alice', age: 25 },
  { name: 'Bob', age: 30 },
];
db.addSheet('Sheet2', initialData);

// Is Sheet Exists
const test2 = db.isSheetExists('Sheet1');
if (test2) {
  console.log("Yes");
} else {
  console.log("No");
}

// Get All Sheet Names
const sheetNames = db.getAllSheetNames();
console.log('Sheet Names:', sheetNames);

// Get Column Datas Number
const value = db.getColumnDatasNumber('name');

// Add Column
db.addColumn('city', 'New York');

// Remove Column
db.removeColumn('city');