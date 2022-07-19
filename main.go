package main

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"
	"github.com/harry1453/go-common-file-dialog/cfd"
	"github.com/harry1453/go-common-file-dialog/cfdutil"
	"github.com/xuri/excelize/v2"
)

func openFileDialog(workbookType string) string {
	result, err := cfdutil.ShowOpenFileDialog(cfd.DialogConfig{
		Title: "Select the " + workbookType + " workbook.",
		Role:  "OpenFileExample",
		FileFilters: []cfd.FileFilter{
			{
				DisplayName: "Excel Files (*.xlsx, *.csv)",
				Pattern:     "*.xlsx;*.csv",
			},
		},
		SelectedFileFilterIndex: 2,
		FileName:                "file.csv",
		DefaultExtension:        "csv",
	})

	if err == cfd.ErrorCancelled {
		fmt.Println("Dialog was cancelled by the user.")
	} else if err != nil {
		fmt.Println(err)
	}
	return result
}

func openFile(workbookPath string) *excelize.File {
	// Excelize read file
	result, err := excelize.OpenFile(workbookPath)
	if err != nil {
		fmt.Println(err)
	}
	
	return result
}

func getHeadings(workbookPath string) []string {
	// Excelize read file
	result, err := excelize.OpenFile(workbookPath)
	if err != nil {
		fmt.Println(err)
	}
	
	// Grabs all the columns from Sheet 1
	// ---TO DO--- get the actual first sheet whatever it's named
	
	// Longform variable declarations 4ever!
	var headings []string 
	
	cols, err := result.Cols("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	
	// For each column, output the value in the first row
	var i int = 1
	for cols.Next() {
		colCell,err := cols.Rows()
		if err != nil {
			fmt.Println(err)
		}
		
		var j string = strconv.Itoa(i) + ": " + colCell[0] + "\n"
		headings = append(headings, j)
		i++
	}
	
	return headings
}

func updateList(workbookPath string, priceColumn int) map[string]string {
	// Excel is not zero-based, but Excelize is so we need to convert our variable
	priceColumn -= 1
	
	var result = openFile(workbookPath)

	rows, err := result.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	
	// Create a map of skus and new prices
	var prices = make(map[string]string)
	
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		prices[row[0]] = row[2]
	}
	
	return prices
}

func importSheet(workbookPath string, prices map[string]string) {
	var result = openFile(workbookPath)
	
	rows, err := result.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	
	for rows.Next() {
		row, err := rows.Columns()
		
		if err != nil {
			fmt.Println(err)
		}
		
		// Grab the part number from column D
		var partNumber string = row[3]
		
		// Clean part number
		if partNumber == "" {
			// ---TO DO--- DO SOMETHING
		}
		// Remove the GRID_ label from part numbers
		if strings.HasPrefix(partNumber, "GRID_") {
			partNumber = strings.TrimPrefix(partNumber, "GRID_")
		}
		// Remove any lowercase suffix
		// Temporarily wrapping this in a check for value since I'm not yet handling blanks
		if partNumber != "" {
			var lastChar string = partNumber[len(partNumber)-1:]
			for _, r := range lastChar {
				if unicode.IsLower(r){
					partNumber = strings.TrimSuffix(partNumber, lastChar)
				}
			}
		}
		
		
		fmt.Println(partNumber)
	}
}

func main() {
	// Get the update sheet and read the column headings
	var uWBPath string = openFileDialog("Update")	
	var headings = getHeadings(uWBPath)
	
	// Ask user to choose the correct column heading
	fmt.Println("Select the column with updated prices")
	for _, val := range headings {
		fmt.Printf(val)
	}
	
	var priceColumn int
	fmt.Scanln(&priceColumn)
			
	// Create a key=>value list (map) of part number and price from the update sheet
	var prices = updateList(uWBPath, priceColumn)
	
	fmt.Println(prices)
	// Get the import sheet and clean part numbers
	var iWBPath string = openFileDialog("Import")
	
	fmt.Println(iWBPath)
	
	importSheet(iWBPath, prices)	
	
	// Match the cleaned part numbers with the key=>value pair list and see if there's an updated price

	// If updated price, insert that into column Q.

	// If no updated price, copy value from column G.

	// If updated price, but price is a string, copy value from column G and insert price into column R

}
