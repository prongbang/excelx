# excelx

[![Go Report Card](https://goreportcard.com/badge/github.com/prongbang/excelx)](https://goreportcard.com/report/github.com/prongbang/excelx)

Convert array struct to XLSX format and Parse XLSX format to array struct with Golang

## Install

```shell
go get github.com/prongbang/excelx
```

## Define struct for Convert

Add `header` for mapping in XLSX header and `no` start with 1 for sort header

```go
type Person struct {
	Name  string `header:"Name" no:"3"`
	Age   int    `header:"Age" no:"1"`
	City  string `header:"City" no:"2"`
	Other string
}
```

## Using for Convert single sheet

```go
m := []MyStruct{
    {"John Doe", 25, "New York", "555"},
    {"Jane Doe", 30, "San Francisco", "555"},
    {"Bob Smith", 22, "Chicago", "555"},
}
file, err := excelx.Convert[MyStruct](m)
```

## Using for Convert multiple sheets

```go
m1 := []MyStruct1{
    {"John Doe", 25, "New York", "555"},
    {"Jane Doe", 30, "San Francisco", "555"},
    {"Bob Smith", 22, "Chicago", "555"},
}
m2 := []MyStruct1{
    {"John Doe", 25, "New York", "555"},
    {"Jane Doe", 30, "San Francisco", "555"},
    {"Bob Smith", 22, "Chicago", "555"},
}
file, err := excelx.Converts(func(file excelx.Xlsx) []excelx.Sheet {
    return []excelx.Sheet{
        {
            Name: "Struct1",
            Exec: func(name string) { excelx.NewSheet(file, name, m1) },
        },
        {
            Name: "Struct1",
            Exec: func(name string) { excelx.NewSheet(file, name, m2) },
        },
    }
})
```

## Save the Excel file to the response writer

- http

```go
err := excelx.ResponseWriter(file, w, "output.xlsx")
```

- fiber

```go
err := excelx.SendStream(ctx, file, "output.xlsx")
```

## Using for Parse

```go
file, _, err := r.FormFile("xlsxfile")
persons, err := excelx.Parse[Person](file)
```

### Example

```go
package main

import (
	"fmt"
	"net/http"

	"github.com/prongbang/excelx"
)

// Define a sample struct
type Person struct {
	Name  string `header:"Name" no:"3"`
	Age   int    `header:"Age" no:"1"`
	City  string `header:"City" no:"2"`
	Other string
}

func generateExcelHandler(w http.ResponseWriter, r *http.Request) {
	// Sample data
	persons := []Person{
		{"John Doe", 25, "New York", "555"},
		{"Jane Doe", 30, "San Francisco", "555"},
		{"Bob Smith", 22, "Chicago", "555"},
	}

	// Create a new Excel file
	file, _ := excelx.Convert[Person](persons)

	// Save the Excel file to the response writer
	err := excelx.ResponseWriter(file, w, "output.xlsx")
	if err != nil {
		fmt.Println("Error writing Excel file to response:", err)
		http.Error(w, "Internal Server Error", http.StatusInternalServerError)
	}
}

func main() {
	http.HandleFunc("/excelx", generateExcelHandler)

	// Start the HTTP server
	fmt.Println("Server listening on :8080...")
	http.ListenAndServe(":8080", nil)
}
```