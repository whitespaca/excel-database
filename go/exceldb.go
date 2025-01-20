package exceldb

import (
	"errors"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

type Row map[string]interface{}

type ExcelDatabase struct {
	FilePath  string
	SheetName string
	Data      []Row
}

func NewExcelDatabase(filePath, sheetName string) (*ExcelDatabase, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		f := excelize.NewFile()
		defaultSheet := f.GetSheetName(f.GetActiveSheetIndex())
		f.SetSheetName(defaultSheet, sheetName)
		if err := f.SaveAs(filePath); err != nil {
			return nil, fmt.Errorf("failed to create new Excel file: %w", err)
		}
	}

	db := &ExcelDatabase{
		FilePath:  filePath,
		SheetName: sheetName,
	}
	err := db.loadData()
	if err != nil {
		return nil, err
	}

	return db, nil
}

func (db *ExcelDatabase) loadData() error {
	f, err := excelize.OpenFile(db.FilePath)
	if err != nil {
		return err
	}
	defer f.Close()

	rows, err := f.GetRows(db.SheetName)
	if err != nil {
		return err
	}

	db.Data = []Row{}
	if len(rows) > 0 {
		headers := rows[0]
		for _, row := range rows[1:] {
			entry := Row{}
			for i, cell := range row {
				if i < len(headers) {
					entry[headers[i]] = cell
				}
			}
			db.Data = append(db.Data, entry)
		}
	}
	return nil
}

func (db *ExcelDatabase) saveData() error {
	f, err := excelize.OpenFile(db.FilePath)
	if err != nil {
		return err
	}
	defer f.Close()

	headers := []string{}
	if len(db.Data) > 0 {
		for key := range db.Data[0] {
			headers = append(headers, key)
		}
	}

	data := [][]string{headers}
	for _, row := range db.Data {
		line := []string{}
		for _, header := range headers {
			if value, ok := row[header]; ok {
				line = append(line, fmt.Sprintf("%v", value))
			} else {
				line = append(line, "")
			}
		}
		data = append(data, line)
	}

	if err := f.DeleteSheet(db.SheetName); err != nil {
		return err
	}
	f.NewSheet(db.SheetName)

	for i, row := range data {
		for j, cell := range row {
			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			f.SetCellValue(db.SheetName, cellName, cell)
		}
	}

	return f.Save()
}

func (db *ExcelDatabase) Select(query Row) ([]Row, error) {
	var result []Row
	for _, row := range db.Data {
		matches := true
		for key, value := range query {
			if row[key] != value {
				matches = false
				break
			}
		}
		if matches {
			result = append(result, row)
		}
	}
	if len(result) == 0 {
		return nil, errors.New("no matching rows found")
	}
	return result, nil
}

func (db *ExcelDatabase) Insert(newRow Row) error {
	db.Data = append(db.Data, newRow)
	return db.saveData()
}

func (db *ExcelDatabase) Update(query Row, updateData Row) error {
	for i, row := range db.Data {
		matches := true
		for key, value := range query {
			if row[key] != value {
				matches = false
				break
			}
		}
		if matches {
			for key, value := range updateData {
				db.Data[i][key] = value
			}
		}
	}
	return db.saveData()
}

func (db *ExcelDatabase) Delete(query Row) error {
	var filtered []Row
	for _, row := range db.Data {
		matches := true
		for key, value := range query {
			if row[key] != value {
				matches = false
				break
			}
		}
		if !matches {
			filtered = append(filtered, row)
		}
	}
	db.Data = filtered
	return db.saveData()
}

func (db *ExcelDatabase) AddColumn(columnName string, defaultValue interface{}) error {
	for i := range db.Data {
		db.Data[i][columnName] = defaultValue
	}
	return db.saveData()
}

func (db *ExcelDatabase) RemoveColumn(columnName string) error {
	for i := range db.Data {
		delete(db.Data[i], columnName)
	}
	return db.saveData()
}

func (db *ExcelDatabase) GetAllSheetNames() ([]string, error) {
	f, err := excelize.OpenFile(db.FilePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return f.GetSheetList(), nil
}

func (db *ExcelDatabase) IsSheetExists(sheetName string) (bool, error) {
	sheets, err := db.GetAllSheetNames()
	if err != nil {
		return false, err
	}
	for _, sheet := range sheets {
		if sheet == sheetName {
			return true, nil
		}
	}
	return false, nil
}
