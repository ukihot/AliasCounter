package main

import (
	"fmt"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

const exportFolderPath string = "./export/"
const importFolderPath string = "./import/"
const importFileName string = "database.xlsx"
const exportFileName string = "Result.xlsx"
const shohinMasterSheet string = "shohin"
const uriageMeisaiSheet string = "uriage"

var shohinRows [][]string
var uriageRows [][]string
var f *excelize.File

// Variable names are unified in Lower CamelCase.
// Prohibit the use of pascal cases and snake cases.

// UriageAnalysis ...
func UriageAnalysis(shohinCode string, shohinName string, kikakuName string, shohinThesaurusChannel chan<- map[string][]string, shohinCodeCounterChannel chan<- map[string]int) {
	shohinThesaurus := map[string][]string{}
	shohinCodeCouner := map[string]int{}

	// 7年分の売上情報を走査
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
	if len(shohinThesaurus) > 0 {
		shohinThesaurusChannel <- shohinThesaurus
	}
	shohinCodeCounterChannel <- shohinCodeCouner
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

// WriteToCell ...
func WriteToCell(s string, r int, c int, t string) {
	tmpCell, _ := excelize.CoordinatesToCellName(c, r)
	f.SetCellValue(s, tmpCell, t)
}

// IsExistSheet ...
func IsExistSheet(str string) bool {
	Sheets := f.GetSheetMap()
	for _, sheet := range Sheets {
		if sheet == str {
			return false
		}
	}
	return true
}

// MakeSheet ...
func MakeSheet(str string, lastRowBySheet map[string]int, JigyosyoName string) {
	if !IsExistSheet(str) {
		return
	}
	// 1行目はヘッダがあるので2行目から
	lastRowBySheet[JigyosyoName] = 2
	f.NewSheet(str)
	// Header creation
	f.SetCellValue(str, "A1", "総使用回数")
	f.SetCellValue(str, "B1", "商品名[商品マスタ]")
	f.SetCellValue(str, "C1", "規格名[商品マスタ]")
}

// OutputToExcel ...
// 事業所別にシート分け
// 売上日 | 担当者 | マスタ商品名 | マスタ規格名 | 売り商品名 | 売り規格名 | 得意先 | 仕入先
func OutputToExcel(shohinThesaurus map[string][]string, shohinCodeCounerThesaurus map[string]int) {
	// Result.xlsxとして結果を保存
	if err := f.SaveAs(exportFolderPath + exportFileName); err != nil {
		fmt.Println(err)
	}
}

// InputFromExcel ...
func InputFromExcel() {
	// SQL実行を保存したExcelファイル、ヘッダ消しといてね。
	database, err := excelize.OpenFile(importFolderPath + importFileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Create an output file.
	f = excelize.NewFile()
	// Retrieves the values of all cells in each row of the product master and stores them as a two-dimensional array.
	shohinRows, err = database.GetRows(shohinMasterSheet)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Retrieves the values of all cells in each row of the sales details and stores them as a two-dimensional array.
	uriageRows, err = database.GetRows(uriageMeisaiSheet)
	if err != nil {
		fmt.Println(err)
		return
	}
}

func main() {
	InputFromExcel()
	// 各ゴルーチンがこの商品コードにはこれだけの別名利用がありましたよと報告するためのチャネル
	shohinThesaurusChannel := make(chan map[string][]string, len(shohinRows))
	shohinCodeCounterChannel := make(chan map[string]int, len(shohinRows))
	wg := new(sync.WaitGroup)
	// 別称利用を格納するシソーラス
	shohinThesaurus := map[string][]string{}
	shohinCodeCounerThesaurus := map[string]int{}

	for _, shohinRow := range shohinRows {
		shohinCode := strings.TrimSpace(shohinRow[0])
		shohinName := strings.TrimSpace(shohinRow[1])
		kikakuName := strings.TrimSpace(shohinRow[2])
		// 商品コード13950件のゴルーチン作成
		wg.Add(1)
		go func() {
			defer wg.Done()
			UriageAnalysis(shohinCode, shohinName, kikakuName, shohinThesaurusChannel, shohinCodeCounterChannel)
		}()
	}
	// ゴルーチンの全完了を検知
	wg.Wait()
	// チャネルクローズ
	close(shohinThesaurusChannel)
	close(shohinCodeCounterChannel)

	// 商品シソーラスチャネルから全データを受信して1つのmapにmerge
	for aliasPage := range shohinThesaurusChannel {
		for shohinCode, aliases := range aliasPage {
			shohinThesaurus[shohinCode] = append(shohinThesaurus[shohinCode], aliases...)
		}
	}

	// 商品コードカウンターチャネルから全データを受信して1つのmapにmerge
	for shohinCodeCounterPage := range shohinCodeCounterChannel {
		for shohinCode, counts := range shohinCodeCounterPage {
			shohinCodeCounerThesaurus[shohinCode] = counts
		}
	}

	OutputToExcel(shohinThesaurus, shohinCodeCounerThesaurus)
}
