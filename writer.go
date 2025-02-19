package goexcel

import (
	"strconv"
	"sync"

	"github.com/xuri/excelize/v2"
)

const MAX_ROW = 1000001

type ExcelWriter struct {
	File         *excelize.File
	streamWriter *excelize.StreamWriter
	header       []interface{}
	currentSheet string
	sheetNum     int
	sheetMaxRow  int
	rowCount     int
	mu           sync.Mutex
}

func Writer() (*ExcelWriter, error) {
	sheetNum := 1
	sheetName := getSheetName(sheetNum)
	f := excelize.NewFile()
	sw, err := f.NewStreamWriter(sheetName)
	if err != nil {
		return nil, err
	}
	return &ExcelWriter{
		File:         f,
		streamWriter: sw,
		currentSheet: sheetName,
		sheetNum:     sheetNum,
		sheetMaxRow:  MAX_ROW,
		rowCount:     0,
	}, nil
}

func getSheetName(sheetNum int) string {
	return "Sheet" + strconv.Itoa(sheetNum)
}

func (w *ExcelWriter) SetSheetMaxRow(sheetMaxRow int) error {
	w.sheetMaxRow = sheetMaxRow
	return nil
}

func (w *ExcelWriter) WriteHeader(headers []interface{}) error {
	if err := w.WriteRow(headers); err != nil {
		return err
	}
	w.header = headers
	return nil
}

func (w *ExcelWriter) WriteRows(data [][]interface{}) error {
	for _, row := range data {
		if err := w.WriteRow(row); err != nil {
			return err
		}
	}
	return nil
}

func (w *ExcelWriter) WriteRow(row []interface{}) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	w.rowCount++
	if w.rowCount > w.sheetMaxRow {
		if err := w.streamWriter.Flush(); err != nil {
			return err
		}
		w.sheetNum++
		newSheetName := getSheetName(w.sheetNum)
		w.File.NewSheet(newSheetName)
		sw, err := w.File.NewStreamWriter(newSheetName)
		if err != nil {
			return err
		}
		w.streamWriter = sw
		// add new sheet header
		w.rowCount = 1
		axis, _ := excelize.CoordinatesToCellName(1, w.rowCount)
		if err = w.streamWriter.SetRow(axis, w.header); err != nil {
			return err
		}
		w.rowCount++
	}

	axis, _ := excelize.CoordinatesToCellName(1, w.rowCount)

	if err := w.streamWriter.SetRow(axis, row); err != nil {
		return err
	}

	return nil
}

func (w *ExcelWriter) Save(filename string) error {
	if err := w.streamWriter.Flush(); err != nil {
		return err
	}
	return w.File.SaveAs(filename)
}

func (w *ExcelWriter) Close() {
	if w.File != nil {
		w.File.Close()
	}
}
