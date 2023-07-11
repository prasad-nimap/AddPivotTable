/* package main

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2")

func processWorkbook(workbook string) *excelize.File {
	f, err := excelize.OpenFile(workbook)
	if err != nil {
		fmt.Print("Failed", err)
		return nil
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Create a new workbook
	newFile := excelize.NewFile()

	// Copy each sheet from the original workbook to the new workbook
	for _, sheet := range f.GetSheetMap() {
		sheetName := sheet
		sheetIndex, _ := f.GetSheetIndex(sheetName)

		// Copy the sheet to the new workbook
		newSheetIndex, err := newFile.NewSheet(sheetName)
		if err != nil {
			fmt.Printf("Failed to create sheet '%s': %v\n", sheetName, err)
			return nil
		}
		err = f.CopySheet(sheetIndex, newSheetIndex)
		if err != nil {
			fmt.Printf("Failed to copy sheet '%s': %v\n", sheetName, err)
			return nil
		}
	}

	// Get the "Interco" sheet from the original workbook
	sheetName := "Interco"
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("Failed to get Interco Sheet Rows:", err)
		return nil
	}

	// Create a new sheet in the new workbook
	newSheetName := "Interco Pivot"
	newSheetIndex, _ := newFile.NewSheet(newSheetName)

	// Copy the rows and data from the "Interco" sheet to the new sheet
	for rowIndex, row := range rows {
		for columnIndex, cellValue := range row {
			cellCoordinates, _ := excelize.CoordinatesToCellName(columnIndex+1, rowIndex+1)
			err := newFile.SetCellValue(newSheetName, cellCoordinates, cellValue)
			if err != nil {
				fmt.Println(err)
				return nil
			}
		}
	}

	// Add the pivot table to the new sheet
	lastRow := len(rows)
	dataRange := fmt.Sprintf("%s!$A$1:$N$%d", newSheetName, lastRow)
	if err := newFile.AddPivotTable(&excelize.PivotTableOptions{
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
		return nil
	}

	// Set the new sheet as active in the new workbook
	newFile.SetActiveSheet(newSheetIndex)

	return newFile
}

func main() {
	args := os.Args[1:]
	if len(args) <= 0 {
		fmt.Println("Usage: go run main.go <workbook>")
		return
	}

	workbook := args[0]
	fmt.Println(workbook)

	newWorkbook := processWorkbook(workbook)
	if newWorkbook == nil {
		fmt.Println("Failed to process workbook")
		return
	}

	// Save the new workbook separately
	newWorkbookPath := "new_" + workbook
	if err := newWorkbook.SaveAs(newWorkbookPath); err != nil {
		fmt.Println("Failed to save new workbook:", err)
		return
	}

	fmt.Println("New workbook saved successfully:", newWorkbookPath)
}
 */