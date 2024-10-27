package excelx

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"mime/multipart"
	"net/http"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

type model[T any] struct {
	Data T
}

type SheetInterface interface {
	GetName() string
	GetData() []any
}

type Sheet struct {
	Name string
	Exec func(name string)
}

type Xlsx struct {
	File *excelize.File
}

type Options struct {
	Options   *excelize.Options
	SheetName string
}

type Response interface {
	Set(key, val string)
	SendStream(stream io.Reader, size ...int) error
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

	defer func() { _ = xlsx.File.Close() }()

	return xlsx.File.GetSheetList()
}

func OpenReader(r io.Reader, opts ...excelize.Options) (Xlsx, error) {
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

	return Xlsx{xlsx}, err
}

func ParseByMultipart[T any](file multipart.File, sheetName ...string) ([]T, error) {
	opts := &Options{}
	if len(sheetName) > 0 {
		opts.SheetName = sheetName[0]
	}
	return Parser[T](file, *opts)
}

func IsEmpty(slice []string) bool {
	// Check if the slice itself is empty
	if len(slice) == 0 {
		return true
	}

	// Iterate over the slice and return false as soon as a non-empty string is found
	for _, s := range slice {
		if s != "" {
			return false
		}
	}
	return true
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
	var xlsx Xlsx
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.File.Close() }()

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
	rows, err := xlsx.File.Rows(sheet)
	if err != nil {
		return []T{}, err
	}

	// Fetch the column names
	_ = rows.Next()
	columns, err := rows.Columns()
	if err != nil {
		return []T{}, err
	}

	// Set the column names as header by index
	header := make([]string, len(columns))
	for i, col := range columns {
		header[i] = col
	}

	// Next the record row
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return []T{}, err
		}

		// Ignore row is empty
		if IsEmpty(cols) {
			continue
		}

		for i, field := range cols {
			structValue := reflect.ValueOf(&record.Data).Elem()
			structField := structValue.FieldByNameFunc(func(name string) bool {
				f, _ := reflect.TypeOf(record.Data).FieldByName(name)
				fieldTag := f.Tag.Get("header")
				head := RemoveDoubleQuote(header[i])
				return fieldTag == fmt.Sprintf("%v", head)
			})

			if structField.IsValid() {
				// Convert the value based on the field kind
				switch structField.Kind() {
				case reflect.Ptr:
					// Handle pointer types
					fieldType := structField.Type()
					elemType := fieldType.Elem()
					ptrValue := reflect.New(elemType)
					switch elemType.Kind() {
					case reflect.String:
						ptrValue.Elem().SetString(field)
					case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
						value, err := strconv.ParseInt(field, 10, 64)
						if err == nil {
							ptrValue.Elem().SetInt(value)
						} else {
							structField.Set(reflect.Zero(fieldType))
							continue
						}
					case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
						value, err := strconv.ParseUint(field, 10, 64)
						if err == nil {
							ptrValue.Elem().SetUint(value)
						} else {
							structField.Set(reflect.Zero(fieldType))
							continue
						}
					case reflect.Float32, reflect.Float64:
						value, err := strconv.ParseFloat(field, 64)
						if err == nil {
							ptrValue.Elem().SetFloat(value)
						} else {
							structField.Set(reflect.Zero(fieldType))
							continue
						}
					case reflect.Bool:
						value, err := strconv.ParseBool(field)
						if err == nil {
							ptrValue.Elem().SetBool(value)
						}
					case reflect.Struct:
						ptrValue.Elem().Set(reflect.ValueOf(field))
					}
					structField.Set(ptrValue)
				default:
					// Handle non-pointer types as before
					switch structField.Kind() {
					case reflect.String:
						structField.SetString(field)
					case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
						value, err := strconv.ParseInt(field, 10, 64)
						if err == nil {
							structField.SetInt(value)
						}
					case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
						value, err := strconv.ParseUint(field, 10, 64)
						if err == nil {
							structField.SetUint(value)
						}
					case reflect.Float32, reflect.Float64:
						value, err := strconv.ParseFloat(field, 64)
						if err == nil {
							structField.SetFloat(value)
						}
					case reflect.Bool:
						value, err := strconv.ParseBool(field)
						if err == nil {
							structField.SetBool(value)
						}
					case reflect.Struct:
						structField.Set(reflect.ValueOf(field))
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
	var xlsx Xlsx
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.File.Close() }()

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
	rows, err := xlsx.File.Rows(sheet)
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
	var xlsx Xlsx
	var err error
	if options != nil {
		xlsx, err = OpenReader(r, *options)
	} else {
		xlsx, err = OpenReader(r)
	}

	defer func() { _ = xlsx.File.Close() }()

	if err != nil {
		return err
	}

	// Iterate through the rows and populate the struct fields
	rows, err := xlsx.File.Rows(sheet)
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

// Convert array struct to Excel format
func Convert[T any](data []T, sheetName ...string) (*Xlsx, error) {
	if len(data) == 0 {
		return nil, errors.New("data is empty")
	}

	// Create a new Excel file
	file := Xlsx{excelize.NewFile()}

	// Create a new sheet
	sheet := "Sheet1"
	if len(sheetName) > 0 {
		sheet = sheetName[0]
	}

	NewSheet(file, sheet, data)

	return &file, nil
}

// Converts array struct to Excel format
func Converts(sheets func(file Xlsx) []Sheet) (Xlsx, error) {
	// Create a new Excel file
	file := Xlsx{excelize.NewFile()}

	// Create a new sheet
	for i, sheet := range sheets(file) {
		if i == 0 {
			_ = file.File.SetSheetName("Sheet1", sheet.Name)
			sheet.Exec(sheet.Name)
		} else {
			sheet.Exec(sheet.Name)
		}
	}

	return file, nil
}

func NewSheet[T any](file Xlsx, sheet string, data []T) {
	_, _ = file.File.NewSheet(sheet)

	// Use reflection to get struct field names and sort them by the "no" tag
	s := reflect.TypeOf(data[0])
	fields := []reflect.StructField{}
	for i := 0; i < s.NumField(); i++ {
		cell := s.Field(i).Tag.Get("header")
		if _, err := strconv.Atoi(s.Field(i).Tag.Get("no")); err == nil && cell != "" {
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
		_ = file.File.SetCellValue(sheet, cell, header)
	}

	// Add data to the sheet using reflection
	for row, md := range data {
		v := reflect.ValueOf(md)
		for col, field := range fields {
			cell := NumberToColName(col+1) + fmt.Sprintf("%d", row+2)
			_ = file.File.SetCellValue(sheet, cell, v.FieldByName(field.Name).Interface())
		}
	}
}

func ResponseWriter(file Xlsx, w http.ResponseWriter, filename string) error {

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// Set the Content-Disposition header to prompt the user to download the file
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))

	// Save the Excel file to the response writer
	return file.File.Write(w)
}

func SendStream[T Response](c T, file Xlsx, filename string) error {
	// Set headers for Excel file download
	c.Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	c.Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))

	// Set up an io.Pipe for efficient memory usage
	pr, pw := io.Pipe()

	// Use a goroutine to write the file to the pipe
	go func() {
		// Ensure the pipe writer closes after writing
		defer func(pw *io.PipeWriter) {
			_ = pw.Close()
		}(pw)

		if file.File != nil {
			if err := file.File.Write(pw); err != nil {
				_ = pw.CloseWithError(err)
			}
		}
	}()

	// Stream the file to the response body with optimal memory usage
	return c.SendStream(pr)
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
