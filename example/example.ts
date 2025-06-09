import { ExcelDatabase } from 'excel-database';

// 1) Initialize database instance (file path and sheet name)
const db = new ExcelDatabase('../data.xlsx', 'Sheet1');

// 2) SELECT rows matching a query
const selectedRows = db.select({ name: 'John Doe' });
console.log('Selected Rows:', selectedRows);

// 3) INSERT a new row
db.insert({ name: 'Jane Doe', age: 30, city: 'New York' });

// 4) UPDATE rows matching a query
db.update({ name: 'Jane Doe' }, { age: 31 });

// 5) DELETE rows matching a query
db.delete({ name: 'John Doe' });

// 6) Get a specific column value
const ageOfJane = db.getColumnValue('name', 'Jane Doe', 'age');
console.log('Jane Doe is', ageOfJane, 'years old');

// 7) ADD a new sheet with initial data
const initialData = [
  { name: 'Alice', age: 25 },
  { name: 'Bob',   age: 30 },
];
db.addSheet('Sheet2', initialData);

// 8) CHECK if a sheet exists
if (db.isSheetExists('Sheet2')) {
  console.log('Sheet2 exists');
} else {
  console.log('Sheet2 does not exist');
}

// 9) LIST all sheet names
console.log('All sheets:', db.getAllSheetNames());

// 10) COUNT non-empty entries in a column
console.log(
  'Number of names:',
  db.getColumnDatasNumber('name')
);

// 11) REMOVE a sheet by name
//    This will delete the sheet and update the workbook accordingly.
//    Throws an Error if the sheet does not exist.
db.removeSheet('Sheet2');
console.log('Removed Sheet2. Current sheets:', db.getAllSheetNames());
