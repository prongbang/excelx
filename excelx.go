package excelx

import (
	"errors"
	"fmt"
	"net/http"
	"reflect"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// NumberToColName converts a column number to an Excel column letter
func NumberToColName(n int) string {
	result := ""
	for n > 0 {
		n--
		result = fmt.Sprintf("%c", 'A'+n%26) + result
		n /= 26
	}
	return result
}

// Convert array struct to excel format
func Convert[T any](data []T, sheetName ...string) (*excelize.File, error) {
	if len(data) == 0 {
		return nil, errors.New("data is empty")
	}

	// Create a new Excel file
	file := excelize.NewFile()

	// Create a new sheet
	sheet := "Sheet1"
	if len(sheetName) > 0 {
		sheet = sheetName[0]
	}
	file.NewSheet(sheet)

	// Use reflection to get struct field names and sort them by the "no" tag
	s := reflect.TypeOf(data[0])
	fields := []reflect.StructField{}
	for i := 0; i < s.NumField(); i++ {
		cell := s.Field(i).Tag.Get("header")
		if _, err := strconv.Atoi(s.Field(i).Tag.Get("no")); err == nil || cell != "" {
			fields = append(fields, s.Field(i))
		}
	}
	sort.Slice(fields, func(i, j int) bool {
		no1, _ := strconv.Atoi(fields[i].Tag.Get("no"))
		no2, _ := strconv.Atoi(fields[j].Tag.Get("no"))
		return no1 < no2
	})

	// Set column headers based on sorted struct fields
	for col, field := range fields {
		header := field.Tag.Get("header")
		cell := NumberToColName(col+1) + "1"
		file.SetCellValue(sheet, cell, header)
	}

	// Add data to the sheet using reflection
	for row, person := range data {
		s := reflect.ValueOf(person)
		for col, field := range fields {
			cell := NumberToColName(col+1) + fmt.Sprintf("%d", row+2)
			file.SetCellValue(sheet, cell, s.FieldByName(field.Name).Interface())
		}
	}

	return file, nil
}

func ResponseWriter(file *excelize.File, w http.ResponseWriter, filename string) error {

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// Set the Content-Disposition header to prompt the user to download the file
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))

	// Save the Excel file to the response writer
	return file.Write(w)
}
