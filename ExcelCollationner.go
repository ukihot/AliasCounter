package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

const exportFolderPath string = "./export/"
const importFolderPath string = "./import/"
const importFileName string = "database.xlsx"
const exportFileName string = "Result.xlsx"
const shohinMasterSheet string = "shohin"
const uriageMeisaiSheet string = "uriage"
const ableCodeSheet string = "移行対象"
const disableCodeSheet string = "DELETE候補"
const checkSheet string = "差異表"

// Variable names are unified in Lower CamelCase.
// Prohibit the use of pascal cases and snake cases.
func main() {
	// Open the data file to be processed.
	database, err := excelize.OpenFile(importFolderPath + importFileName)
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
	result.SetCellValue(ableCodeSheet, "B1", "商品コード")
	result.SetCellValue(ableCodeSheet, "C1", "商品マスタ")
	result.SetCellValue(ableCodeSheet, "D1", "売上明細")
	result.SetCellValue(disableCodeSheet, "A1", "商品コード")
	result.SetCellValue(disableCodeSheet, "B1", "商品名 / 規格")
	// The product code in the sales invoice is linked to the product master, but the product name in the sales invoice can change in many ways.
	// There are probably more than 80% of the product codes in the product master that are not needed.
	i, j := 2, 2
	shohinCodeCounter := map[string]int{}
	cellAdressMemory := map[string]string{}
	// Import data file format
	// ./data/shohin.sql
	// CAUTION: No header required.
	for _, shohinRow := range shohinRows {
		shohinCode := shohinRow[0]
		for _, uriageRow := range uriageRows { // If you find that a product code in the product master is used in sales, save it in result.
			if uriageRow[0] == shohinCode {
				if _, isThere := shohinCodeCounter[shohinCode]; isThere { // If the product code is already known, add the product name.
					// Increment the number of uses
					shohinCodeCounter[shohinCode]++
					result.SetCellValue(ableCodeSheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
					// Update the cell number to add the product name.
					_, r, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					shohinNameCell, _ := excelize.CoordinatesToCellName(shohinCodeCounter[shohinCode]+2, r)
					standardNameCell, _ := excelize.CoordinatesToCellName(shohinCodeCounter[shohinCode]+2, r+1)
					dateCell, _ := excelize.CoordinatesToCellName(shohinCodeCounter[shohinCode]+2, r+2)
					tmpCell, _ := excelize.CoordinatesToCellName(1, r+1)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinCodeCounter[shohinCode])
					tmpCell, _ = excelize.CoordinatesToCellName(1, r+2)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinCodeCounter[shohinCode])
					result.SetCellValue(ableCodeSheet, shohinNameCell, uriageRow[1])
					result.SetCellValue(ableCodeSheet, standardNameCell, uriageRow[2])
					result.SetCellValue(ableCodeSheet, dateCell, uriageRow[4])
				} else { //If it is a new product code, fill in the valid sheet.
					cellAdress, _ := excelize.CoordinatesToCellName(1, i)
					// Initialize the number of uses
					result.SetCellValue(ableCodeSheet, cellAdress, 1)
					shohinCodeCounter[shohinCode] = 1
					cellAdressMemory[shohinCode] = cellAdress
					// 商品コード
					c, _, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					tmpCell, _ := excelize.CoordinatesToCellName(c, i+1)
					result.SetCellValue(ableCodeSheet, tmpCell, 1)
					tmpCell, _ = excelize.CoordinatesToCellName(c, i+2)
					result.SetCellValue(ableCodeSheet, tmpCell, 1)
					tmpCell, _ = excelize.CoordinatesToCellName(c+1, i)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinCode)
					tmpCell, _ = excelize.CoordinatesToCellName(c+1, i+1)
					result.SetCellValue(ableCodeSheet, tmpCell, "規格")
					tmpCell, _ = excelize.CoordinatesToCellName(c+1, i+2)
					result.SetCellValue(ableCodeSheet, tmpCell, "登録日")
					//商品名
					tmpCell, _ = excelize.CoordinatesToCellName(c+2, i)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinRow[1])
					//規格名
					tmpCell, _ = excelize.CoordinatesToCellName(c+2, i+1)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinRow[2])
					//登録日
					tmpCell, _ = excelize.CoordinatesToCellName(c+2, i+2)
					result.SetCellValue(ableCodeSheet, tmpCell, shohinRow[4])
					// Increment the line number
					i = i + 3
				}
			}
		}
		if _, isThere := shohinCodeCounter[shohinCode]; !isThere {
			cellAdress, _ := excelize.CoordinatesToCellName(1, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinCode)
			cellAdress, _ = excelize.CoordinatesToCellName(2, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinRow[1])
			cellAdress, _ = excelize.CoordinatesToCellName(3, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinRow[2])
			cellAdress, _ = excelize.CoordinatesToCellName(3, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinRow[4])
			j++
		}
	}
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil { // Save the file for output.
		fmt.Println(err)
	}
}
