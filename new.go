/*
	package main

import (

	"fmt"

	"github.com/xuri/excelize/v2"

)

	func checkError(err error) {
		if err != nil {
			fmt.Print("Something went wrong", err)
			return
		}
	}

	func main() {
		// Open source file
		f, err := excelize.OpenFile("cashcopy.xlsx")
		checkError(err)

		// Get names of sheets from the workbook
		sheetList := f.GetSheetList()

		// Print all the sheets names
		for index, sheetname := range sheetList {
			fmt.Println(sheetname, " has ", index)
		}

		// Catch the name of the sheet to be copied
		sn := "May"

		// Get the destination sheet name
		dsn := "May"

		// Get all the rows
		rows, err := f.GetRows(sn)
		checkError(err)

		// New file to store the copied sheet
		d := excelize.NewFile()

		_, err := f.CopySheet(f, 1)
		checkError(err)

		// save the new file
		err = d.SaveAs("result.xlsx")
		checkError(err)

		fmt.Println("Copied")
	}
*/
package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Open the source workbook
	srcFile, err := excelize.OpenFile("cashcopy.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Get the source sheet name
	srcSheetName := "May"

	// Get all the source sheet's rows and columns
	rows, err := srcFile.GetRows(srcSheetName, excelize.Options{RawCellValue: true})
	if err != nil {
		log.Fatal(err)
	}

	// Create a new workbook for the destination
	dstFile := excelize.NewFile()

	// Get the destination sheet name
	dstSheetName := "CopiedSheet" // Set the desired name for the destination sheet

	// Create the destination sheet in the destination workbook
	_, err = dstFile.NewSheet(dstSheetName)

	// Copy rows from the source sheet to the destination sheet
	for rowIndex, row := range rows {
		for columnIndex, cellValue := range row {
			// Get the cell coordinates (e.g., A1, B2, etc.)
			cellCoords, err := excelize.CoordinatesToCellName(columnIndex+1, rowIndex+1)
			if err != nil {
				log.Fatal(err)
			}

			// for style
			cellStyleId, _ := srcFile.GetCellStyle(srcSheetName, cellCoords)
			ws, err := srcFile.workSheetReader(sheet)
			ws.SheetData.Row[rowIndex].C[columnIndex].S
			cellStyle := srcFile.GetSt
			fmt.Println(cellStyle)
			err = dstFile.SetCellStyle(dstSheetName, cellCoords, cellCoords, cellStyle)

			// for formula
			formula, err := srcFile.GetCellFormula(srcSheetName, cellCoords)
			if formula == "" {
				// Set the cell value in the destination sheet
				err = dstFile.SetCellValue(dstSheetName, cellCoords, cellValue)
				if err != nil {
					log.Fatal(err)
				}
			} else {
				err = dstFile.SetCellFormula(dstSheetName, cellCoords, formula)
				if err != nil {
					log.Fatal(err)
				}
			}

		}
	}

	// Save the destination workbook to a new file
	err = dstFile.SaveAs("result.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Sheet copied and saved successfully!")
}
