package goexcel

import (
	"reflect"
	"slices"
	"strconv"
)

// RowDecode reflect excel row to struct
func RowDecode(row, header []string, output interface{}) error {
	v := reflect.ValueOf(output).Elem()
	t := v.Type()
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		excelTag := field.Tag.Get("excel")
		if excelTag == "" {
			continue
		}
		fieldValue := v.Field(i)
		kind := fieldValue.Type().Kind()
		j := slices.Index(header, excelTag)
		if j == -1 {
			continue
		}
		value := ""
		if len(row)-1 >= j {
			value = row[j]
		}
		if kind == reflect.Ptr {
			elemType := field.Type.Elem()
			switch elemType.Kind() {
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				intVal, err := strconv.ParseInt(value, 10, 64)
				if err != nil {
					return err
				}
				ptr := reflect.New(elemType)
				ptr.Elem().SetInt(intVal)
				fieldValue.Set(ptr)
			case reflect.Float32, reflect.Float64:
				floatValue, err := strconv.ParseFloat(value, 64)
				if err != nil {
					return err
				}
				ptr := reflect.New(elemType)
				ptr.Elem().SetFloat(floatValue)
				fieldValue.Set(ptr)
			default:
				ptr := reflect.New(elemType)
				ptr.Elem().SetString(value)
				fieldValue.Set(ptr)
			}
		} else {
			switch kind {
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				intValue, err := strconv.ParseInt(value, 10, 64)
				if err != nil {
					return err
				}
				fieldValue.SetInt(intValue)
			case reflect.Float32, reflect.Float64:
				floatValue, err := strconv.ParseFloat(value, 64)
				if err != nil {
					return err
				}
				fieldValue.SetFloat(floatValue)
			default:
				fieldValue.SetString(value)
			}
		}
	}

	return nil
}
