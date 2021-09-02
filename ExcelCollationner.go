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
const nullExpression string = "« NULL »"

func main() {
	// 処理対象データファイルを開く
	database, err := excelize.OpenFile(inputFolderPath + inputFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// 出力用ファイルを作成
	result := excelize.NewFile()
	result.NewSheet(ableCodeSheet)
	result.NewSheet(disableCodeSheet)
	result.DeleteSheet("Sheet1")
	result.SetActiveSheet(0)

	// 商品マスタの行ごとにすべてのセルの値を取得し、2次元配列として格納
	shohinRows, err := database.GetRows(shohinMasterSheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	// 売上明細の行ごとにすべてのセルの値を取得し、2次元配列として格納
	uriageRows, err := database.GetRows(uriageMeisaiSheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	// ヘッダー作成
	result.SetCellValue(ableCodeSheet, "A1", "有効商品コード")
	result.SetCellValue(ableCodeSheet, "B1", "使用回数")
	result.SetCellValue(disableCodeSheet, "A1", "不要商品コード")

	// 売上明細の商品コードは商品マスタに紐づいているが、売上明細側では商品名が多様に変化している
	// どのような商品名に変化をしているのか、
	// 商品マスタ内で不要な商品コードがおそらく８割強あるため、洗い出す
	i, j := 1, 1
	shohinCodeCounter := map[string]int{}
	for _, shohinRow := range shohinRows {
		shohinCode := shohinRow[0]
		for _, uriageRow := range uriageRows {
			// 商品マスタ内の商品コードが売上で使われていることが分かったらresultに保存
			if uriageRow[0] == shohinCode {
				// 保存項目は商品コード、商品名（売上明細）、使用回数
				cellAdress, _ := excelize.CoordinatesToCellName(1, i)
				result.SetCellValue(ableCodeSheet, cellAdress, shohinCode)
				if _, isThere := shohinCodeCounter[shohinCode]; isThere {
					shohinCodeCounter[shohinCode]++
				} else {
					shohinCodeCounter[shohinCode] = 0
				}
				i++
			} else {
				cellAdress, _ := excelize.CoordinatesToCellName(1, j)
				result.SetCellValue(disableCodeSheet, cellAdress, shohinCode)
				j++
			}
		}
	}
	// 出力用ファイルを保存
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil {
		fmt.Println(err)
	}
}
