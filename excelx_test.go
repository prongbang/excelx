package excelx_test

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
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

type Simple struct {
	A *string  `header:"A"`
	B *float64 `header:"B"`
	C *float64 `header:"C"`
	D *float64 `header:"D"`
	E *float64 `header:"E"`
	F *float64 `header:"F"`
	G *float64 `header:"G"`
	H *float64 `header:"H"`
	I *float64 `header:"I"`
	J *float64 `header:"J"`
	K *float64 `header:"K"`
	L *float64 `header:"L"`
	M *float64 `header:"M"`
	N *float64 `header:"N"`
	O *float64 `header:"O"`
	P *float64 `header:"P"`
	Q *float64 `header:"Q"`
	R *float64 `header:"R"`
	S *float64 `header:"S"`
	T *float64 `header:"T"`
	U *float64 `header:"U"`
	V *float64 `header:"V"`
	W *float64 `header:"W"`
	X *float64 `header:"X"`
	Y *float64 `header:"Y"`
	Z *float64 `header:"Z"`
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

func TestParser(t *testing.T) {
	// Given
	data, err := os.ReadFile("simple.xlsx")
	reader := bytes.NewReader(data)
	expect := `{"A":"A1","B":2,"C":3,"D":4,"E":5,"F":5,"G":null,"H":7,"I":8,"J":9,"K":10,"L":11.5,"M":12,"N":5,"O":14.5,"P":null,"Q":null,"R":2.3,"S":1.4,"T":1.4,"U":1.4,"V":1.4,"W":1.4,"X":4,"Y":null,"Z":null}`

	// When
	simple, err := excelx.Parser[Simple](reader)

	// Then
	if err != nil {
		t.Error("Not error", err)
	}
	fmt.Println(len(simple))
	for _, s := range simple {
		b, _ := json.Marshal(s)
		if string(b) != expect {
			t.Error("Parse error", string(b))
		}
	}
}
