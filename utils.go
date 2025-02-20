package goexcel

import (
	"github.com/xuri/excelize/v2"
)

func GetSheetRowCount(f *excelize.File, sheet string) int64 {
	rows, err := f.Rows(sheet)
	if err != nil {
		return 0
	}
	var count int64
	for rows.Next() {
		count++
	}
	return count
}
