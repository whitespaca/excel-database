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

#### Remove Sheet

```typescript
// Delete the sheet named 'OldData'
db.removeSheet('OldData');

// Verify removal
console.log(db.getAllSheetNames());
```

#### Is Sheet Exists

```typescript
const value = db.isSheetExists('sheetName');
```

#### Get All Sheet Names

```typescript
const value = db.getAllSheetNames();
```

#### Get Column Datas Number

```typescript
const value = db.getColumnDatasNumber('columnName');
```

#### Add Column

```typescript
db.addColumn('newColumn', 'defaultValue');
```

#### Remove Column

```typescript
db.removeColumn('columnName');
```
