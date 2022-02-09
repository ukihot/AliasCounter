package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"strings"

	"sync"
)

const exportFolderPath string = "./export/"
const importFolderPath string = "./import/"
const importFileName string = "database.xlsx"
const exportFileName string = "Result.xlsx"
const shohinMasterSheet string = "shohin"
const uriageMeisaiSheet string = "uriage"
const aliasSummarySheet string = "ALIAS_SUMMARY_SHEET"
var wg sync.WaitGroup
// Variable names are unified in Lower CamelCase.
// Prohibit the use of pascal cases and snake cases.
func main() {
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
	// チャネル作成
	shohinThesaurusChannel := make(chan map[string][]string)
	shohinCodeCounerChannel := make(chan map[string]int)

	for _, shohinRow := range shohinRows {
		shohinCode := strings.TrimSpace(shohinRow[0])
		shohinName := strings.TrimSpace(shohinRow[1])
		kikakuName := strings.TrimSpace(shohinRow[2])
		// 商品コード13950件のゴルーチン作成
		wg.Add(1)
		go aliasCounter(result, uriageRows, shohinCode, shohinName, kikakuName, shohinThesaurusChannel, shohinCodeCounerChannel)
	}
	wg.Wait()
	// ゴルーチンの結果データを受信
	select {
	case shohinThesaurus := <-shohinThesaurusChannel:
		for shohinCode, aliases := range shohinThesaurus {
			println(shohinCode, aliases)
		}
	}

	// Result.xlsxとして結果を保存
	if err := result.SaveAs(exportFolderPath + exportFileName); err != nil {
		fmt.Println(err)
	}
}

func aliasCounter(result *excelize.File, uriageRows [][]string, shohinCode string, shohinName string, kikakuName string, shohinThesaurusChannel chan map[string][]string, shohinCodeCounterChannel chan map[string]int) {
	shohinThesaurus := map[string][]string{}
	shohinCodeCouner := map[string]int{}
	for _, uriageRow := range uriageRows {
		uriageShohinCode := strings.TrimSpace(uriageRow[0])
		uriageShohinName := strings.TrimSpace(uriageRow[1])
		// uriageKikakuName := strings.TrimSpace(uriageRow[2])
		// uriageTantoName := strings.TrimSpace(uriageRow[3])
		if shohinCode != uriageShohinCode {
			continue
		}
		// 売上明細の中に当該商品コードを発見したのでカウント
		if _, found := shohinCodeCouner[shohinCode]; found {
			shohinCodeCouner[shohinCode]++
		} else {
			shohinCodeCouner[shohinCode] = 1
		}
		// 商品シソーラスに未登録であれば登録
		if (shohinName != uriageShohinName) && !SliceContains(shohinThesaurus[shohinCode], uriageShohinName) {
			shohinThesaurus[shohinCode] = append(shohinThesaurus[shohinCode], uriageShohinName)
		}
	}
	shohinThesaurusChannel <- shohinThesaurus
	shohinCodeCounterChannel <- shohinCodeCouner
	defer wg.Done()
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

func toCell(ef *excelize.File, s string, r int, c int, t string) {
	tmpCell, _ := excelize.CoordinatesToCellName(c, r)
	ef.SetCellValue(s, tmpCell, t)
}
