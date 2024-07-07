package excelx

import (
	"errors"
	"fmt"
	"io"
	"log"
	"mime/multipart"
	"net/http"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type model[T any] struct {
	Data T
}

type Options struct {
	Options   *excelize.Options
	SheetName string
}

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

// Convert string number to int
func toInt(s string) int {
	value, err := strconv.Atoi(s)
	if err != nil {
		return 0
	}
	return value
}

func GetSheetList(r io.Reader, opts ...excelize.Options) []string {
	// Open the XLSX file
	xlsx, err := OpenReader(r, opts...)
	if err != nil {
		log.Println(err)
		return []string{}
	}

	defer func() { _ = xlsx.Close() }()

	return xlsx.GetSheetList()
}

func OpenReader(r io.Reader, opts ...excelize.Options) (*excelize.File, error) {
	var options *excelize.Options
	if len(opts) > 0 {
		options = &opts[0]
	}
	var xlsx *excelize.File
	var err error
	if options != nil {
		xlsx, err = excelize.OpenReader(r, *options)
	} else {
		xlsx, err = excelize.OpenReader(r)
	}

	return xlsx, err
}

func ParseByMultipart[T any](file multipart.File, sheetName ...string) ([]T, error) {
	opts := &Options{}
	if len(sheetName) > 0 {
		opts.SheetName = sheetName[0]
	}
	return Parser[T](file, *opts)
}

// Parser excel format to array struct
func Parser[T any](r io.Reader, opts ...Options) ([]T, error) {
	// Set sheet name
	sheet := "Sheet1"
	var options *excelize.Options
	if len(opts) > 0 {
		if opts[0].Options != nil {
			options = opts[0].Options
		}
		if opts[0].SheetName != "" {
			sheet = opts[0].SheetName
		}
	}

	// Open the XLSX file
	var xlsx *excelize.File
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.Close() }()

	if err != nil {
		return []T{}, err
	}

	// Extract header information from the struct tags
	headerMap := make(map[int]string)
	record := model[T]{}
	modelType := reflect.TypeOf(record.Data)
	for i := 0; i < modelType.NumField(); i++ {
		field := modelType.Field(i)
		header := field.Tag.Get("header")
		no := field.Tag.Get("no")
		if header != "" && no != "" {
			headerMap[toInt(no)] = header
		}
	}

	// Iterate through the rows and populate the struct fields
	var records []T
	rows, err := xlsx.Rows(sheet)
	if err != nil {
		return []T{}, err
	}

	// Skip the header row
	rows.Next()

	// Next the record row
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return []T{}, err
		}

		for i, col := range cols {
			fieldName := headerMap[i+1]
			if fieldName != "" {
				structValue := reflect.ValueOf(&record.Data).Elem()
				field := structValue.FieldByNameFunc(func(name string) bool {
					f, _ := reflect.TypeOf(record.Data).FieldByName(name)
					fieldTag := f.Tag.Get("header")
					head := RemoveDoubleQuote(fieldName)
					return fieldTag == fmt.Sprintf("%v", head)
				})
				if field.IsValid() {
					// Convert the value based on the field kind
					switch field.Kind() {
					case reflect.String:
						field.SetString(col)
					case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
						value, err := strconv.ParseInt(col, 10, 64)
						if err == nil {
							field.SetInt(value)
						}
					case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
						value, err := strconv.ParseUint(col, 10, 64)
						if err == nil {
							field.SetUint(value)
						}
					case reflect.Float32, reflect.Float64:
						value, err := strconv.ParseFloat(col, 64)
						if err == nil {
							field.SetFloat(value)
						}
					case reflect.Bool:
						value, err := strconv.ParseBool(col)
						if err == nil {
							field.SetBool(value)
						}
					case reflect.Struct:
						// Assuming the time is represented in the format "2006-01-02 15:04:05"
						value, err := time.Parse("2006-01-02 15:04:05", col)
						if err == nil {
							field.Set(reflect.ValueOf(value))
						}
					}
				}
			}
		}

		records = append(records, record.Data)
	}

	return records, nil
}

// ParserString excel format to array struct
func ParserString[T any](r io.Reader, opts ...Options) ([]T, error) {
	// Set sheet name
	sheet := "Sheet1"
	var options *excelize.Options
	if len(opts) > 0 {
		if opts[0].Options != nil {
			options = opts[0].Options
		}
		if opts[0].SheetName != "" {
			sheet = opts[0].SheetName
		}
	}

	// Open the XLSX file
	var xlsx *excelize.File
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.Close() }()

	if err != nil {
		return []T{}, err
	}

	// Extract header information from the struct tags
	headerMap := make(map[int]string)
	record := model[T]{}
	modelType := reflect.TypeOf(record.Data)
	for i := 0; i < modelType.NumField(); i++ {
		field := modelType.Field(i)
		header := field.Tag.Get("header")
		no := field.Tag.Get("no")
		if header != "" && no != "" {
			headerMap[toInt(no)] = header
		}
	}

	// Iterate through the rows and populate the struct fields
	var records []T
	rows, err := xlsx.Rows(sheet)
	if err != nil {
		return []T{}, err
	}

	// Skip the header row
	rows.Next()

	// Next the record row
	for rows.Next() {
		cols, e := rows.Columns()
		if e != nil {
			return []T{}, e
		}

		for i, col := range cols {
			fieldName := headerMap[i+1]
			if fieldName != "" {
				structValue := reflect.ValueOf(&record.Data).Elem()
				field := structValue.FieldByNameFunc(func(name string) bool {
					f, _ := reflect.TypeOf(record.Data).FieldByName(name)
					fieldTag := f.Tag.Get("header")
					head := RemoveDoubleQuote(fieldName)
					return fieldTag == fmt.Sprintf("%v", head)
				})
				if field.IsValid() {
					field.SetString(col)
				}
			}
		}

		records = append(records, record.Data)
	}

	return records, nil
}

// ParserFunc excel format to array struct
func ParserFunc(r io.Reader, onRecord func([]string) error, opts ...Options) error {
	// Set sheet name
	sheet := "Sheet1"
	var options *excelize.Options
	if len(opts) > 0 {
		if opts[0].Options != nil {
			options = opts[0].Options
		}
		if opts[0].SheetName != "" {
			sheet = opts[0].SheetName
		}
	}

	// Open the XLSX file
	var xlsx *excelize.File
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.Close() }()

	if err != nil {
		return err
	}

	// Iterate through the rows and populate the struct fields
	rows, err := xlsx.Rows(sheet)
	if err != nil {
		return err
	}

	// Next the record row
	for rows.Next() {
		cols, e := rows.Columns()
		if e != nil {
			return e
		}

		er := onRecord(cols)
		if er != nil {
			return er
		}
	}

	return nil
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
	_, _ = file.NewSheet(sheet)

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
		_ = file.SetCellValue(sheet, cell, header)
	}

	// Add data to the sheet using reflection
	for row, model := range data {
		s := reflect.ValueOf(model)
		for col, field := range fields {
			cell := NumberToColName(col+1) + fmt.Sprintf("%d", row+2)
			_ = file.SetCellValue(sheet, cell, s.FieldByName(field.Name).Interface())
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

func RequestFile(r *http.Request, name string) (multipart.File, *multipart.FileHeader, error) {
	return r.FormFile(name)
}

// RemoveDoubleQuote remove double quote (") and clear unicode from text
func RemoveDoubleQuote(text string) string {
	text = ClearUnicode(text)
	first := strings.Index(text, "\"")
	last := strings.LastIndex(text, "\"")
	if first == 0 && last == (len(text)-1) {
		text = text[1:last]
	}
	return text
}

// ClearUnicode clear unicode from text
func ClearUnicode(text string) string {
	regex := regexp.MustCompile("^\ufeff")
	result := regex.ReplaceAllString(text, "")
	return strings.TrimSpace(result)
}
