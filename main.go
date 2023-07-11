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

	// Sheet1 already exists...
	index, err := f.NewSheet("copysheet.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	_, err := f.CopySheet(1, index)
} */
package main

import (
    "fmt"
    // "os"

    "github.com/xuri/excelize/v2"
)

func columnNumberToLetter(columnNumber int) string {
    columnLetter := ""
    for columnNumber > 0 {
        remainder := (columnNumber - 1) % 26
        columnLetter = string('A'+remainder) + columnLetter
        columnNumber = (columnNumber - 1) / 26
    }
    return columnLetter
}

func other() {
    /* args := os.Args[1:] // Skip the first argument as it contains the program name
    if len(args) < 2 {
        fmt.Println("Usage: go run main.go <source workbook> <destination workbook>")
        return
    }

    sourceWorkbook := args[0]
    destinationWorkbook := args[1]
    fmt.Println("Source workbook:", sourceWorkbook)
    fmt.Println("Destination workbook:", destinationWorkbook) */

    // Open the source workbook
    sourceFile, err := excelize.OpenFile("copysheet.xlsx")
    if err != nil {
        fmt.Println("Failed to open source workbook:", err)
        return
    }
    defer func() {
        if err := sourceFile.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    // Create a new destination workbook
    // destFile := excelize.NewFile()

    // Copy sheets from the source workbook to the destination workbook
    sourceSheets := []string{"Settings", "InterCO Comments"}
    for _, sourceSheet := range sourceSheets {
        // Read data from the source workbook
        rows, err := sourceFile.GetRows(sourceSheet)
        if err != nil {
            fmt.Printf("Failed to get rows from source sheet '%s': %v\n", sourceSheet, err)
            continue
        }

        // Create a new sheet in the destination workbook
        destSheet := sourceSheet
        destFile.NewSheet(destSheet)

        // Write data to the destination workbook
        for rowIndex, row := range rows {
            for colIndex, cellValue := range row {
                // Set the cell value in the destination workbook
                colLetter := columnNumberToLetter(colIndex + 1)
                destFile.SetCellValue(destSheet, colLetter+fmt.Sprint(rowIndex+1), cellValue)
            }
        }
    }

    // Save the destination workbook
    if err := destFile.SaveAs("cash.xlsx"); err != nil {
        fmt.Println("Failed to save destination workbook:", err)
        return
    }

    fmt.Println("Sheets copied and pasted successfully.")
}