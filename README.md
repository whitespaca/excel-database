# ExcelDatabase

![npm](https://img.shields.io/npm/v/excel-database) ![npm](https://img.shields.io/npm/dt/excel-database)

A lightweight library for managing Excel files as databases in Node.js, built with TypeScript.

## Features

- **CRUD Operations**: Perform create, read, update, and delete operations on Excel files.
- **Column Lookup**: Search rows by column values and fetch specific column values.
- **TypeScript Support**: Fully typed for better development experience.

## Installation

Install the library via [npm](https://npmjs.org/package/excel-database):

```bash
npm install excel-database
```

## Usage

### Import and Initialize

```typescript
import { ExcelDatabase } from 'excel-database';

const db = new ExcelDatabase('path/to/your/file.xlsx');

// Default sheet name is 'Sheet1'. You can specify a custom sheet name:
const dbWithCustomSheet = new ExcelDatabase('path/to/your/file.xlsx', 'CustomSheetName');
```

### CRUD Operations

#### Select

```typescript
const results = db.select({ columnName: 'value' });
console.log(results);
```

#### Insert

```typescript
db.insert({ columnName: 'newValue', anotherColumn: 123 });
```

#### Update

```typescript
db.update({ columnName: 'value' }, { columnName: 'updatedValue' });
```

#### Delete

```typescript
db.delete({ columnName: 'value' });
```

#### Get Column Value

```typescript
const value = db.getColumnValue('searchColumn', 'searchValue', 'targetColumn');
console.log(value);
```

#### Add Sheet
```typescript
// Add New Sheet With Data (Recommand)
const initialData = [
  { value: 'value', value_1: 1 },
  { value2: 'value2', value2_1: 2 },
];

db.addSheet('NewSheet', initialData);

// Add New Empty Sheet
db.addSheet('EmptySheet');
```

#### Is Sheet Exists
```typescript
const value = db.isSheetExists('sheetName');
```

#### Get All Sheet Names
```typescript
const value = db.getAllSheetNames();
```

## Example
```typescript
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
if (value) {
  console.log("Yes");
} else {
  console.log("No");
}

// Get All Sheet Names
const sheetNames = db.getAllSheetNames();
console.log('Sheet Names:', sheetNames);
```

## Contributing

Please submit issues or pull requests via [GitHub](https://github.com/whitespaca/excel-database).

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.