package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const exportFolderPath string = "./export/"
const inputFolderPath string = "./data/"
const inputFileName string = "database.xlsx"
const exportFileName string = "Result.xlsx"
const shohinMasterSheet string = "shohin"
const uriageMeisaiSheet string = "uriage"
const ableCodeSheet string = "有効商品コード"
const disableCodeSheet string = "無効商品コード"

// Variable names are unified in Lower CamelCase.
// Prohibit the use of pascal cases and snake cases.
func main() {
	// Open the data file to be processed.
	database, err := excelize.OpenFile(inputFolderPath + inputFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Create an output file.
	result := excelize.NewFile()
	result.NewSheet(ableCodeSheet)
	result.NewSheet(disableCodeSheet)
	result.DeleteSheet("Sheet1")
	result.SetActiveSheet(0)
	// Retrieves the values of all cells in each row of the product master and stores them as a two-dimensional array.
	shohinRows, err := database.GetRows(shohinMasterSheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Retrieves the values of all cells in each row of the sales details and stores them as a two-dimensional array.
	uriageRows, err := database.GetRows(uriageMeisaiSheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Header creation
	result.SetCellValue(ableCodeSheet, "A1", "使用回数")
	result.SetCellValue(ableCodeSheet, "B1", "商品名 / 規格")
	result.SetCellValue(disableCodeSheet, "A1", "商品コード")
	result.SetCellValue(disableCodeSheet, "B1", "商品名 / 規格")
	// The product code in the sales invoice is linked to the product master, but the product name in the sales invoice can change in many ways.
	// There are probably more than 80% of the product codes in the product master that are not needed.
	i, j := 2, 2
	shohinCodeCounter := map[string]int{}
	cellAdressMemory := map[string]string{}
	// Import data file format
	// [売上明細シート]商品コード,商品名,規格,登録ユーザID,登録日
	// [商品マスタシート]商品コード,商品名,規格,登録ユーザID,登録日
	// Export data file format
	// [有効シート]使用回数,商品マスタの商品コード,商品マスタの<商品名+規格>,売上明細の<商品名+規格>
	// [無効シート]商品マスタの商品コード,商品マスタの商品名
	for _, shohinRow := range shohinRows {
		shohinCode := shohinRow[0]
		for _, uriageRow := range uriageRows { // If you find that a product code in the product master is used in sales, save it in result.
			if uriageRow[0] == shohinCode {
				if _, isThere := shohinCodeCounter[shohinCode]; isThere { // 既出の商品コードだった場合商品名を連ねる
					// Increment the number of uses
					shohinCodeCounter[shohinCode]++
					result.SetCellValue(ableCodeSheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
					// Update the cell number to add the product name.
					c, r, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					shohinNameCell, _ := excelize.CoordinatesToCellName(c+shohinCodeCounter[shohinCode], r)
					// Combining product name and standard name
					uriageDetail := uriageRow[1]
					if uriageRow[2] != "" {
						uriageDetail = uriageRow[1] + " / " + uriageRow[2]
					}
					result.SetCellValue(ableCodeSheet, shohinNameCell, uriageDetail)
				} else { //If it is a new product code, fill in the valid sheet.
					cellAdress, _ := excelize.CoordinatesToCellName(1, i)
					shohinCodeCounter[shohinCode] = 1
					cellAdressMemory[shohinCode] = cellAdress
					// Initialize the number of uses
					result.SetCellValue(ableCodeSheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
					// Update the cell number to add the product name.
					c, _, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					tmpCell, _ := excelize.CoordinatesToCellName(c, i)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinCode)
					// Update the cell number to enter the product name + standard.
					tmpCell, _ = excelize.CoordinatesToCellName(c+1, i)
					// Combining product name and standard name
					shohinDetail := shohinRow[1]
					if shohinRow[2] != "" {
						shohinDetail = shohinRow[1] + " / " + shohinRow[2]
					}
					result.SetCellValue(ableCodeSheet, tmpCell, shohinDetail)
					// Increment the line number
					i++
				}
			}
		}
		if _, isThere := shohinCodeCounter[shohinCode]; !isThere {
			cellAdress, _ := excelize.CoordinatesToCellName(1, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinCode)
			cellAdress, _ = excelize.CoordinatesToCellName(2, j)
			// Combining product name and standard name
			shohinDetail := shohinRow[1]
			if shohinRow[2] != "" {
				shohinDetail = shohinRow[1] + " / " + shohinRow[2]
			}
			result.SetCellValue(disableCodeSheet, cellAdress, shohinDetail)
			j++
		}
	}
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil { // Save the file for output.
		fmt.Println(err)
	}
}
