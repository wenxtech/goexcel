package goexcel

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func GetSheetRowCount(f *excelize.File, sheetName string) (int64, error) {
	// get sheet dimension info
	dimension, err := f.GetSheetDimension(sheetName)
	if err != nil {
		return 0, err
	}

	if dimension == "" {
		return 0, nil
	}
	// split dimension info. eg"A1:B2" -> ["A1", "B2"]
	parts := strings.Split(dimension, ":")
	if len(parts) != 2 {
		return 0, fmt.Errorf("invalid dimension format: %s", dimension)
	}
	// match row number B2 -> 2
	match := regexp.MustCompile(`^[A-Za-z]+(\d+)$`).FindStringSubmatch(parts[1])
	if len(match) < 2 {
		return 0, fmt.Errorf("failed to parse row from coordinate: %s", parts[1])
	}

	rows, err := strconv.ParseInt(match[1], 10, 64)
	if err != nil {
		return 0, fmt.Errorf("invalid row number: %s", match[1])
	}

	return rows, nil
}
