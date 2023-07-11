/* package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("cash.xlsx")
	if err != nil {
		fmt.Print("Failed", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sheetName := "Apr"
	// rows, err1 := f.GetRows(sheetName)
	// if err1 != nil {
	// 	fmt.Println("Failed to get Interco Sheet Rows:", err1)
	// 	return
	// }

	row := 7
	cellName, _ := excelize.CoordinatesToCellName(row, 272)
	fmt.Println(cell)
	cellFormula, _ := f.GetCellFormula(sheetName, cellName)

	if cellFormula == "" {
		fmt.Println("NO FORMULA")
	}

	// for _, row := range rows {
	// 	for _, cell := range row {
	// 		fmt.Println(cell)
	// 	}
	// }

	// lastRow := len(rows)

	// dataRange := fmt.Sprintf("%s!$A$1:$N$%d", sheetName, lastRow)

	// if err := f.AddPivotTable(&excelize.PivotTableOptions{
	// 	DataRange:       dataRange,
	// 	PivotTableRange: "Interco Pivot!$A$2:$C$2000",
	// 	Rows: []excelize.PivotTableField{
	// 		{Data: "Row Field", DefaultSubtotal: true},
	// 		{Data: "Interco Entity"},
	// 	},
	// 	Columns: []excelize.PivotTableField{},
	// 	Data: []excelize.PivotTableField{
	// 		{Data: "Sep", Name: "Sep Total", Subtotal: "Sum"},
	// 		{Data: "YTD", Name: "YTD Total", Subtotal: "Sum"},
	// 	},
	// 	RowGrandTotals: true,
	// 	ColGrandTotals: true,
	// 	ShowDrill:      true,
	// 	ShowRowHeaders: true,
	// 	ShowColHeaders: true,
	// 	ShowLastColumn: true,
	// }); err != nil {
	// 	fmt.Println(err)
	// }
	f.SaveAs("book1.xlsx")
}

// package main

// import (
// 	"fmt"
// 	// "math/rand"
// 	"os"

// 	"github.com/xuri/excelize/v2"
// )

// func main() {
//  workbook := os.Args[1]
//  f, err := excelize.OpenFile(workbook)
//  if err != nil {
//      fmt.Print("Failed", err)
//  }
//  defer func() {
//      if err := f.Close(); err != nil {
//          fmt.Println(err)
//      }
//  }()
//  sheetName := "Interco"
//  rows, err := f.GetRows(sheetName)
//  if err != nil {
//      fmt.Println("Failed to get Interco Sheet Rows:", err)
//      return
//  }

//  lastRow := len(rows)

//  dataRange := fmt.Sprintf("%s!$A$1:$N$%d", sheetName, lastRow)

//  if err := f.AddPivotTable(&excelize.PivotTableOptions{
//      DataRange:       dataRange,
//      PivotTableRange: "Interco Pivot!$A$2:$C$2000",
//      Rows: []excelize.PivotTableField{
//          {Data: "Row Field", DefaultSubtotal: true},
//          {Data: "Interco Entity"},
//      },
//      Columns: []excelize.PivotTableField{},
//      Data: []excelize.PivotTableField{
//          {Data: "Sep", Name: "Sep Total", Subtotal: "Sum"},
//          {Data: "YTD", Name: "YTD Total", Subtotal: "Sum"},
//      },
//      RowGrandTotals: true,
//      ColGrandTotals: true,
//      ShowDrill:      true,
//      ShowRowHeaders: true,
//      ShowColHeaders: true,
//      ShowLastColumn: true,
//  }); err != nil {
//      fmt.Println(err)
//  }
//  f.SaveAs("test.xlsx")
// }

// func processWorkbook(workbook string) *excelize.File {
// 	f, err := excelize.OpenFile(workbook)
// 	if err != nil {
// 		fmt.Print("Failed", err)
// 		return nil
// 	}
// 	defer func() {
// 		if err := f.Close(); err != nil {
// 			fmt.Println(err)
// 		}
// 	}()

// 	// sheetName := "May"
// 	// rows, err := f.GetRows(sheetName)
// 	// if err != nil {
// 	// 	fmt.Println("Failed to get Interco Sheet Rows:", err)
// 	// 	return nil
// 	// }
// 	// lastRow := len(rows)
// 	// f.CopySheet(2, 3)

// 	f.SaveAs("text.xlsx")

// 	// dataRange := fmt.Sprintf("%s!$A$1:$N$%d", sheetName, lastRow)

// 	// if err := f.AddPivotTable(&excelize.PivotTableOptions{
// 	// 	DataRange:       dataRange,
// 	// 	PivotTableRange: "Interco Pivot!$A$2:$C$2000",
// 	// 	Rows: []excelize.PivotTableField{
// 	// 		{Data: "Row Field", DefaultSubtotal: true},
// 	// 		{Data: "Interco Entity"},
// 	// 	},
// 	// 	Columns: []excelize.PivotTableField{},
// 	// 	Data: []excelize.PivotTableField{
// 	// 		{Data: "Sep", Name: "Sep Total", Subtotal: "Sum"},
// 	// 		{Data: "YTD", Name: "YTD Total", Subtotal: "Sum"},
// 	// 	},
// 	// 	RowGrandTotals: true,
// 	// 	ColGrandTotals: true,
// 	// 	ShowDrill:      true,
// 	// 	ShowRowHeaders: true,
// 	// 	ShowColHeaders: true,
// 	// 	ShowLastColumn: true,
// 	// }); err != nil {
// 	// 	fmt.Println(err)
// 	// }

// 	return f
// }

// func main() {

// 	args := os.Args[1:] // Skip the first argument as it contains the program name
// 	if len(args) <= 0 {
// 		fmt.Println("Usage: go run main.go <workbook>")
// 		return
// 	}

// 	workbook := args[0]
// 	fmt.Println(workbook)

// 	f := processWorkbook(workbook)
// 	if f == nil {
// 		fmt.Println("Failed to process workbook")
// 		return
// 	}

// 	f.SaveAs(workbook)
// }
 */