package excelx_test

import (
	"testing"

	"github.com/prongbang/excelx"
)

// Define a sample struct
type Person struct {
	Name  string `header:"Name" no:"3"`
	Age   int    `header:"Age" no:"1"`
	City  string `header:"City" no:"2"`
	Other string
}

func TestNumberToColName(t *testing.T) {
	// Given
	column := 10

	// When
	actual := excelx.NumberToColName(column)

	// Then
	if actual != "J" {
		t.Error("Error", actual)
	}
}

func TestConvertIsNotEmpty(t *testing.T) {
	// Given
	persons := []Person{
		{"John Doe", 25, "New York", "555"},
		{"Jane Doe", 30, "San Francisco", "555"},
		{"Bob Smith", 22, "Chicago", "555"},
	}

	// When
	_, err := excelx.Convert[Person](persons)

	// Then
	if err != nil {
		t.Error("Error", err)
	}
}

func TestConvertIsEmpty(t *testing.T) {
	// Given
	persons := []Person{}

	// When
	_, err := excelx.Convert[Person](persons)

	// Then
	if err == nil {
		t.Error("Not error", err)
	}
}
