package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"strings"

	"strconv"

	"github.com/cheggaaa/pb/v3"
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
	// [1.4万行]×[57万行]の2重Loopが遅すぎるため、商品コード1レコード読むたびに進捗率を更新
	bar := pb.StartNew(13950)
	bar.SetMaxWidth(80)
	// SQL実行を保存したExcelファイル、ヘッダ消しといてね。
	database, err := excelize.OpenFile(importFolderPath + importFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Create an output file.
	result := excelize.NewFile()
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
	// The product code in the sales invoice is linked to the product master, but the product name in the sales invoice can change in many ways.
	// There are probably more than 80% of the product codes in the product master that are not needed.
	shohinCodeCounter := map[string]int{}
	shohinThesaurus := map[string][]string{}
	kikakuThesaurus := map[string][]string{}
	uriageShohinNameCounter := map[string]int{}
	uriageKikakuNameCounter := map[string]int{}
	cellRowMemory := map[string]int{}
	sheetMemory := map[string]string{}
	lastRowBySheet := map[string]int{}
	var shohinCode string
	var shohinName string
	var kikakuName string
	var uriageShohinCode string
	var uriageShohinName string
	var uriageKikakuName string
	var uriageTantoName string
	// var uriageDate string
	// var tokuisakiName string
	// var shiiresakiName string
	// var jigyoshoName string

	// Import data file format
	// ./data/select.sql
	// CAUTION: No header required.
	for _, shohinRow := range shohinRows {
		shohinCode = strings.TrimSpace(shohinRow[0])
		shohinName = strings.TrimSpace(shohinRow[1])
		kikakuName = strings.TrimSpace(shohinRow[2])
		// If you find that a product code in the product master is used in sales, save it in result.
		for _, uriageRow := range uriageRows {
			uriageShohinCode = strings.TrimSpace(uriageRow[0])
			uriageShohinName = strings.TrimSpace(uriageRow[1])
			uriageKikakuName = strings.TrimSpace(uriageRow[2])
			uriageTantoName = strings.TrimSpace(uriageRow[3])
			// uriageDate = strings.TrimSpace(uriageRow[4])
			// tokuisakiName = strings.TrimSpace(uriageRow[5])
			// shiiresakiName = strings.TrimSpace(uriageRow[6])
			// jigyoshoName = strings.TrimSpace(uriageRow[7])
			// 商品コードが合致しないものは無視
			if shohinCode != uriageShohinCode {
				continue
			}
			// 売上担当者ごとにシート分ける
			MakeSheet(result, uriageTantoName, lastRowBySheet, uriageShohinName)
			if _, isThere := shohinCodeCounter[shohinCode]; !isThere { // 探していた商品コード発見１回目の処理
				sheetMemory[shohinCode] = uriageTantoName
				// Increment the line number
				shohinCodeCounter[shohinCode] = 1
				// 1列目：商品コード総使用回数
				tmpCell, _ := excelize.CoordinatesToCellName(1, lastRowBySheet[uriageTantoName])
				result.SetCellValue(uriageTantoName, tmpCell, shohinCodeCounter[shohinCode])
				// 2列目：商品名
				tmpCell, _ = excelize.CoordinatesToCellName(2, lastRowBySheet[uriageTantoName])
				result.SetCellValue(uriageTantoName, tmpCell, shohinName)
				// 3列目：規格名
				tmpCell, _ = excelize.CoordinatesToCellName(3, lastRowBySheet[uriageTantoName])
				result.SetCellValue(uriageTantoName, tmpCell, kikakuName)
				// マスタと違う商品名で販売しているか
				RegisterThesaurus(shohinThesaurus, shohinCode, uriageShohinName, uriageShohinNameCounter, kikakuThesaurus, kikakuName, uriageKikakuName, uriageKikakuNameCounter, shohinName)
				// セルの行を記憶して次行へ
				cellRowMemory[shohinCode] = lastRowBySheet[uriageTantoName]
				lastRowBySheet[uriageTantoName]++
			} else {
				// 商品コード使用回数をインクリメント
				shohinCodeCounter[shohinCode]++
				tmpCell, _ := excelize.CoordinatesToCellName(10, cellRowMemory[shohinCode])
				result.SetCellValue(uriageTantoName, tmpCell, "hoge")
				// マスタと違う商品名で販売しているか
				RegisterThesaurus(shohinThesaurus, shohinCode, uriageShohinName, uriageShohinNameCounter, kikakuThesaurus, kikakuName, uriageKikakuName, uriageKikakuNameCounter, shohinName)
			}
		}
		// 1列目：商品コード総使用回数
		tmpCell, _ := excelize.CoordinatesToCellName(1, cellRowMemory[shohinCode])
		result.SetCellValue(uriageTantoName, tmpCell, shohinCodeCounter[shohinCode])
		// 1商品コードについて全売上情報を参照した状態
		//
		// 商品シソーラスがあれば商品コードに紐づく行にて横展開
		if shohinAliases, isThere := shohinThesaurus[shohinCode]; isThere {
			for index, shohinAlias := range shohinAliases {
				tmpCell, _ := excelize.CoordinatesToCellName(4+index, cellRowMemory[shohinCode])
				result.SetCellValue(sheetMemory[shohinCode], tmpCell, shohinAlias+"("+strconv.Itoa(uriageShohinNameCounter[shohinAlias])+")")
			}
		}
		bar.Increment()
	}
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil {
		fmt.Println(err)
	}
	result.DeleteSheet("Sheet1")
	bar.Finish()
}

// SliceContains ...
func SliceContains(arr []string, str string) bool {
	for _, v := range arr {
		if v == str {
			return true
		}
	}
	return false
}

// IsExistSheet ...
func IsExistSheet(result *excelize.File, str string) bool {
	Sheets := result.GetSheetMap()
	for _, sheet := range Sheets {
		if sheet == str {
			return false
		}
	}
	return true
}

// MakeSheet ...
func MakeSheet(result *excelize.File, str string, lastRowBySheet map[string]int, uriageTantoName string) {
	if !IsExistSheet(result, str) {
		return
	}
	// 1行目はヘッダがあるので2行目から
	lastRowBySheet[uriageTantoName] = 2
	result.NewSheet(str)
	// Header creation
	result.SetCellValue(str, "A1", "総使用回数")
	result.SetCellValue(str, "B1", "商品名[商品マスタ]")
	result.SetCellValue(str, "C1", "規格名[商品マスタ]")
}

// RegisterThesaurus ...
func RegisterThesaurus(shohinThesaurus map[string][]string, shohinCode string, uriageShohinName string, uriageShohinNameCounter map[string]int, kikakuThesaurus map[string][]string, kikakuName string, uriageKikakuName string, uriageKikakuNameCounter map[string]int, shohinName string) {
	// シソーラスに未登録の売り商品名
	if shohinName != uriageShohinName && !SliceContains(shohinThesaurus[shohinCode], uriageShohinName) {
		shohinThesaurus[shohinCode] = append(shohinThesaurus[shohinCode], uriageShohinName)
		uriageShohinNameCounter[uriageShohinName] = 1
	} else {
		// 登録済みなら別称利用回数をインクリメント
		uriageShohinNameCounter[uriageShohinName]++
	}
	// シソーラスに未登録の売り規格名
	if kikakuName != uriageShohinName && !SliceContains(kikakuThesaurus[kikakuName], uriageKikakuName) {
		kikakuThesaurus[kikakuName] = append(kikakuThesaurus[kikakuName], uriageKikakuName)
		uriageKikakuNameCounter[uriageKikakuName] = 1
	} else {
		// 登録済みなら別称利用回数をインクリメント
		uriageKikakuNameCounter[uriageKikakuName]++
	}
}
