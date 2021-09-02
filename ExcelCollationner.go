package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const exportFolderPath string = "./export/"
const inputFolderPath string = "./data/"
const inputFileName string = "database.xlsx"
const shohinMasterSheetName string = "shohin"

func main() {
	f, err := excelize.OpenFile(inputFolderPath + inputFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err := f.GetRows(shohinMasterSheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
}
