package goexcel

import (
	"errors"
	"fmt"

	"github.com/xuri/excelize/v2"
)

type ExcelReader struct {
	File               *excelize.File
	sheetNames         []string
	ableSheets         []string
	ableSheetLen       int
	currentSheetIndex  int
	currentSheetName   string
	currentSheetHeader []string
	currentRows        *excelize.Rows
	currentRow         []string
	currentRowIndex    int
}

// Reader get a new reader
func Reader(filePath string, opts ...excelize.Options) (*ExcelReader, error) {
	f, err := excelize.OpenFile(filePath, opts...)
	if err != nil {
		return nil, err
	}

	// get all sheet name
	sheetMap := f.GetSheetMap()
	sheetNames := make([]string, 0, len(sheetMap))
	for _, name := range sheetMap {
		sheetNames = append(sheetNames, name)
	}

	er := &ExcelReader{
		File:         f,
		sheetNames:   sheetNames,
		ableSheets:   sheetNames,
		ableSheetLen: len(sheetNames),
	}

	// init first sheet
	if len(sheetNames) > 0 {
		if err = er.initSheet(0); err != nil {
			f.Close()
			return nil, err
		}
	}

	return er, nil
}

func (er *ExcelReader) AbleSheet(s []string) {
	er.ableSheets = s
	er.ableSheetLen = len(er.ableSheets)
	er.initSheet(0)
}

func (er *ExcelReader) GetHeader() []string {
	return er.currentSheetHeader
}

// initSheet init current sheet
func (er *ExcelReader) initSheet(sheetIndex int) error {
	if sheetIndex >= er.ableSheetLen {
		return nil
	}

	// when end close reader
	if er.currentRows != nil {
		er.currentRows.Close()
	}

	// get rows reader
	rows, err := er.File.Rows(er.ableSheets[sheetIndex])
	if err != nil {
		return err
	}
	er.currentSheetName = er.ableSheets[sheetIndex]

	er.currentSheetIndex = sheetIndex
	er.currentRows = rows

	// get header
	if er.currentRows.Next() {
		er.currentSheetHeader, err = er.currentRows.Columns()
		if err != nil {
			return err
		}
	}
	er.currentRowIndex = 1

	return nil
}

// Next move to next line
func (er *ExcelReader) Next() (err error) {
	for {
		// check is end of current sheet
		if er.currentRows == nil {
			return ErrEnd
		}
		// try to get next row
		if er.currentRows.Next() {
			er.currentRowIndex++
			er.currentRow, err = er.currentRows.Columns()
			return err
		}

		// check is end of all sheets
		if er.currentSheetIndex+1 >= er.ableSheetLen {
			er.currentRows = nil
			return ErrEnd
		}

		// init next sheet
		if err = er.initSheet(er.currentSheetIndex + 1); err != nil {
			return err
		}
	}
}

// NextScan move to next line and scan row value
func (er *ExcelReader) NextScan(row interface{}) (err error) {
	err = er.Next()
	if err != nil {
		if errors.Is(err, ErrEnd) {
			return nil
		}
		return err
	}
	if err = RowDecode(er.currentRow, er.currentSheetHeader, row); err != nil {
		return err
	}
	return nil
}

// Value get current row value
func (er *ExcelReader) Value() []string {
	return er.currentRow
}

// CurrRowIndex get current row index
func (er *ExcelReader) CurrRowIndex() int {
	return er.currentRowIndex
}

// CurrSheetName get current sheet name
func (er *ExcelReader) CurrSheetName() string {
	return er.currentSheetName
}

func (er *ExcelReader) ErrorInfo(err error) string {
	return fmt.Sprintf("Sheet: %s, Row: %d, Error: %s", er.currentSheetName, er.currentRowIndex, err)
}

// IsEnd check reader is end
func (er *ExcelReader) IsEnd() bool {
	if er.currentRows == nil {
		return true
	}
	return false
}

// Close reader`
func (er *ExcelReader) Close() {
	if er.currentRows != nil {
		er.currentRows.Close()
	}
	if er.File != nil {
		er.File.Close()
	}
}

// RowCount get total row count
func (er *ExcelReader) RowCount() int64 {
	var count int64
	for _, sheet := range er.ableSheets {
		count += GetSheetRowCount(er.File, sheet)
	}
	return count
}

// DataCount get total data count without header line
func (er *ExcelReader) DataCount() int64 {
	return er.RowCount() - int64(len(er.ableSheets))
}
