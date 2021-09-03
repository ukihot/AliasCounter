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
	i, j := 2, 2
	shohinCodeCounter := map[string]int{}
	cellAdressMemory := map[string]string{}
	// インポートデータファイルの形式
	// [売上明細シート]商品コード,商品名,規格,登録ユーザID,登録日
	// [商品マスタシート]商品コード,商品名,規格,登録ユーザID,登録日
	// エクスポートデータファイルの形式
	// [有効シート]使用回数,商品マスタの商品コード,商品マスタの<商品名+規格>,売上明細の<商品名+規格>
	// [無効シート]商品マスタの商品コード,商品マスタの商品名
	for _, shohinRow := range shohinRows {
		shohinCode := shohinRow[0]
		for _, uriageRow := range uriageRows { // 商品マスタ内の商品コードが売上で使われていることが分かったらresultに保存
			if uriageRow[0] == shohinCode {
				if _, isThere := shohinCodeCounter[shohinCode]; isThere { // 既出の商品コードだった場合商品名を連ねる
					// 使用回数のインクリメント
					shohinCodeCounter[shohinCode]++
					result.SetCellValue(ableCodeSheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
					// 商品名を追記するセル番地の更新
					c, r, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					shohinNameCell, _ := excelize.CoordinatesToCellName(c+shohinCodeCounter[shohinCode], r)
					// 商品名と規格を結合
					uriageDetail := uriageRow[1]
					if uriageRow[2] != "" {
						uriageDetail = uriageRow[1] + " / " + uriageRow[2]
					}
					// 商品名+規格を記入
					result.SetCellValue(ableCodeSheet, shohinNameCell, uriageDetail)
				} else { //新規の商品コードだった場合有効シートに記入
					cellAdress, _ := excelize.CoordinatesToCellName(1, i)
					shohinCodeCounter[shohinCode] = 1
					cellAdressMemory[shohinCode] = cellAdress
					// 使用回数の初期値
					result.SetCellValue(ableCodeSheet, cellAdressMemory[shohinCode], shohinCodeCounter[shohinCode])
					// 商品コードを記入するセル番地の更新
					c, _, _ := excelize.CellNameToCoordinates(cellAdressMemory[shohinCode])
					tmpCell, _ := excelize.CoordinatesToCellName(c, i)
					// 商品コードを記入
					result.SetCellValue(ableCodeSheet, tmpCell, shohinCode)
					// 商品名+規格を記入するセル番地の更新
					tmpCell, _ = excelize.CoordinatesToCellName(c+1, i)
					// 商品名と規格を結合
					shohinDetail := shohinRow[1]
					if shohinRow[2] != "" {
						shohinDetail = shohinRow[1] + " / " + shohinRow[2]
					}
					// 商品名+規格を記入
					result.SetCellValue(ableCodeSheet, tmpCell, shohinDetail)
					// 行番号をインクリメント
					i++
				}
			}
		}
		if _, isThere := shohinCodeCounter[shohinCode]; !isThere {
			cellAdress, _ := excelize.CoordinatesToCellName(1, j)
			result.SetCellValue(disableCodeSheet, cellAdress, shohinCode)
			cellAdress, _ = excelize.CoordinatesToCellName(2, j)
			// 商品名と規格を結合
			shohinDetail := shohinRow[1]
			if shohinRow[2] != "" {
				shohinDetail = shohinRow[1] + " / " + shohinRow[2]
			}
			result.SetCellValue(disableCodeSheet, cellAdress, shohinDetail)
			j++
		}
	}
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil { // 出力用ファイルを保存
		fmt.Println(err)
	}
}
