package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const exportFolderPath string = "./export/"
const inputFolderPath string = "./data/"
const inputFileName string = "database.xlsx"

func main() {
	f := excelize.NewFile()

	if err := f.SaveAs("サンプル.xlsx"); err != nil {
		fmt.Println(err)
	}
}
