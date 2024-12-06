# ExcelDatabase

A lightweight library for managing Excel files as databases in Node.js, built with TypeScript.

## Features

- **CRUD Operations**: Perform create, read, update, and delete operations on Excel files.
- **Column Lookup**: Search rows by column values and fetch specific column values.
- **TypeScript Support**: Fully typed for better development experience.

## Installation

Install the library via npm:

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

## Contributing

Contributions are welcome! Please submit issues or pull requests via [GitHub](https://github.com/whitespaca/excel-database).

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
