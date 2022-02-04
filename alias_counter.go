package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"strings"

	"strconv"
)

const exportFolderPath string = "./export/"
const importFolderPath string = "./import/"
const importFileName string = "database.xlsx"
const exportFileName string = "Result.xlsx"
const shohinMasterSheet string = "shohin"
const uriageMeisaiSheet string = "uriage"
const aliasSummarySheet string = "ALIAS_SUMMARY_SHEET"

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
	result.NewSheet(aliasSummarySheet)
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
	result.SetCellValue(aliasSummarySheet, "A1", "使用回数")
	result.SetCellValue(aliasSummarySheet, "B1", "商品コード")
	result.SetCellValue(aliasSummarySheet, "C1", "商品マスタ")
	result.SetCellValue(aliasSummarySheet, "D1", "売上明細")
	// The product code in the sales invoice is linked to the product master, but the product name in the sales invoice can change in many ways.
	// There are probably more than 80% of the product codes in the product master that are not needed.
	i := 2
	shohinCodeCounter := map[string]int{}
	thesaurus := map[string][]string{}
	uriageShohinNameCounter := map[string]int{}
	cellAdressMemory := map[string]string{}

	// Import data file format
	// ./data/select.sql
	// CAUTION: No header required.
	for _, shohinRow := range shohinRows {
		shohinCode := strings.TrimSpace(shohinRow[0])
		shohinName := strings.TrimSpace(shohinRow[1])
		for _, uriageRow := range uriageRows { // If you find that a product code in the product master is used in sales, save it in result.
			uriageShohinCode := strings.TrimSpace(uriageRow[0])
			uriageShohinName := strings.TrimSpace(uriageRow[1])
			uriageKikakuName := strings.TrimSpace(uriageRow[2])
			uriageTantoName := strings.TrimSpace(uriageRow[3])
			// uriageDate := strings.TrimSpace(uriageRow[4])
			if shohinCode != uriageShohinCode {
				continue
			}
			if _, isThere := shohinCodeCounter[shohinCode]; isThere { // If the product code is already known, add the product name.
				// Increment the number of uses
				shohinCodeCounter[shohinCode]++
				result.SetCellValue(aliasSummarySheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
				if uriageShohinName != shohinName {
					//登録済みでない
					if !sliceContains(thesaurus[shohinName], uriageShohinName) {
						thesaurus[shohinName] = append(thesaurus[shohinName], uriageShohinName)
						uriageShohinNameCounter[uriageShohinName] = 1
					} else {
						// 登録済みなら別称利用回数をインクリメント
						uriageShohinNameCounter[uriageShohinName]++
					}
				}
			} else { // Only works the first time.
				tmpCell, _ := excelize.CoordinatesToCellName(1, i)
				cellAdressMemory[shohinCode] = tmpCell
				result.SetCellValue(aliasSummarySheet, tmpCell, 1)
				tmpCell, _ = excelize.CoordinatesToCellName(2, i)
				result.SetCellValue(aliasSummarySheet, tmpCell, shohinCode)
				tmpCell, _ = excelize.CoordinatesToCellName(2, i+1)
				result.SetCellValue(aliasSummarySheet, tmpCell, uriageTantoName)
				//商品名
				tmpCell, _ = excelize.CoordinatesToCellName(3, i)
				result.SetCellValue(aliasSummarySheet, tmpCell, shohinName)
				//規格名
				tmpCell, _ = excelize.CoordinatesToCellName(3, i+1)
				result.SetCellValue(aliasSummarySheet, tmpCell, uriageKikakuName)
				// Increment the line number
				i = i + 2
				shohinCodeCounter[shohinCode] = 1
			}
		}
		//
		// 1商品コードについて全売上情報を参照した状態
		//
		if alias, isThere := thesaurus[shohinName]; isThere {
			for seed, uj := range alias {
				cellAdress, _ := excelize.CoordinatesToCellName(4+seed, i-2)
				result.SetCellValue(aliasSummarySheet, cellAdress, strconv.Itoa(uriageShohinNameCounter[uj])+":"+uj)
			}
		}
	}
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil { // Save the file for output.
		fmt.Println(err)
	}
}

func sliceContains(arr []string, str string) bool {
	for _, v := range arr {
		if v == str {
			return true
		}
	}
	return false
}
