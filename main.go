package main

import (
	"fmt"
	// "math/rand"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("Balance Sheet Activity 202305-70 test.xlsx")
	if err != nil {
		fmt.Print("Failed", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sheetName := "Interco"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("Failed to get Interco Sheet Rows:", err)
		return
	}

	lastRow := len(rows)

	dataRange := fmt.Sprintf("%s!$A$1:$N$%d", sheetName, lastRow)

	if err := f.AddPivotTable(&excelize.PivotTableOptions{
		DataRange:       dataRange,
		PivotTableRange: "Interco Pivot!$A$2:$C$2000",
		Rows: []excelize.PivotTableField{
			{Data: "Row Field", DefaultSubtotal: true},
			{Data: "Interco Entity"},
		},
		Columns: []excelize.PivotTableField{},
		Data: []excelize.PivotTableField{
			{Data: "Sep", Name: "Sep Total", Subtotal: "Sum"},
			{Data: "YTD", Name: "YTD Total", Subtotal: "Sum"},
		},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
	}
	f.SaveAs("book1.xlsx")
}
