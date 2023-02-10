package main

import (
	"context"
	"fmt"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"strconv"
	"strings"
	"time"
)

// 読み込むスプレットシートID
var spreadsheetID1 = "1SAwRc11TMl9fc8243Es0HvL8ZK3SHa6nEyTprCBC6Bk"

// 書き込むスプレットシートID
var spreadsheetID2 = "1jo4DVvChI5sNXFlyUhqF8oSmfTVyiryhiJy6KHkd_xw"

func GetValuesInSpreadSheet(srv *sheets.Service, spreadsheetID, range_ string) (*sheets.Spreadsheet, error) {
	resp, err := srv.Spreadsheets.Get(spreadsheetID).IncludeGridData(true).Ranges(range_).Do()
	if err != nil {
		return nil, err
	}
	return resp, nil
}

// TODO: スプレットシートのタイトルを変更できるようにする。
// TODO: 交通費の情報がいくつあるか指定しなくても持ってこれて、書き込めるようにしたい。
func main() {

	now := time.Now()
	// 請求日
	billdate := fmt.Sprintf("%d/%d/%d", now.Year(), int(now.Month()), now.Day())
	// 給料日
	payDate := fmt.Sprintf("%d/%d/15", now.Year(), int(now.Month()))

	// コンストラクタ?を作成
	srv, err := sheets.NewService(context.TODO(), option.WithCredentialsFile("credentials/secret.json"))
	if err != nil {
		log.Fatal(err)
	}

	// シートとセルを指定、範囲で指定する場合は A1:B6 のようにする
	readRange1 := "大村 2023/02!C9"      // 勤務時間
	readRange2 := "大村 2023/02!A20:E28" // 交通費情報 (多めにセルを指定しておく)

	// 勤務時間を取得
	wrkHr, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange1)
	if err != nil {
		log.Fatal(err)
	}

	// 交通費の情報を取得
	trsptExpnss, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange2)
	if err != nil {
		log.Fatal(err)
	}

	// スプレットシートから読み込んだセルの値、今回だと勤務時間
	workHour := wrkHr.Sheets[0].Data[0].RowData[0].Values[0].FormattedValue
	// 25:50 のように、: が含まれる文字列を小数に変換するために : から . へ置き換える
	workHour = strings.Replace(workHour, ":", ".", 1)
	// 小数として扱いたいので、string 型を float64 型に変換
	workHours, err := strconv.ParseFloat(workHour, 64)
	if err != nil {
		log.Fatalln(err)
	}

	var (
		station []string
		count   []string
		price   []string
	)

	for _, s := range trsptExpnss.Sheets {
		for _, row := range s.Data[0].RowData {
			station = append(station, fmt.Sprintf("%v %v %v", row.Values[0].FormattedValue, row.Values[1].FormattedValue,
				row.Values[2].FormattedValue))
			count = append(count, row.Values[3].FormattedValue)
			price = append(price, row.Values[4].FormattedValue)
		}
	}
	fmt.Println(station)

	//sprdSht := &sheets.BatchUpdateSpreadsheetRequest{
	//	Properties: &sheets.SpreadsheetProperties{
	//		Title: "New Spreadsheet", // スプレッドシートの名前
	//	},
	//}

	// 更新範囲と更新値の指定
	valueRange1 := "N4"
	values1 := [][]interface{}{
		{billdate},
	}
	valueRange2 := "M15"
	values2 := [][]interface{}{
		{payDate},
	}
	valueRange3 := "J18"
	values3 := [][]interface{}{
		{workHours},
	}
	valueRange4 := "A20:A28"
	values4 := [][]interface{}{
		{station[0]},
		{station[1]},
		{station[2]},
		{station[3]},
		{station[4]},
		{station[5]},
		{station[6]},
		{station[7]},
		{station[8]},
	}
	valueRange5 := "J20:J28"
	values5 := [][]interface{}{
		{count[0]},
		{count[1]},
		{count[2]},
		{count[3]},
		{count[4]},
		{count[5]},
		{count[6]},
		{count[7]},
		{count[8]},
	}
	valueRange6 := "L20:L28"
	values6 := [][]interface{}{
		{price[0]},
		{price[1]},
		{price[2]},
		{price[3]},
		{price[4]},
		{price[5]},
		{price[6]},
		{price[7]},
		{price[8]},
	}

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data: []*sheets.ValueRange{
			{
				Range:          valueRange1,
				MajorDimension: "ROWS",
				Values:         values1,
			},
			{
				Range:          valueRange2,
				MajorDimension: "ROWS",
				Values:         values2,
			},
			{
				Range:          valueRange3,
				MajorDimension: "ROWS",
				Values:         values3,
			},
			{
				Range:          valueRange4,
				MajorDimension: "ROWS",
				Values:         values4,
			},
			{
				Range:          valueRange5,
				MajorDimension: "ROWS",
				Values:         values5,
			},
			{
				Range:          valueRange6,
				MajorDimension: "ROWS",
				Values:         values6,
			},
		},
	}

	_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, rb).Do()
	//_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, sprdSht).Do()
	if err != nil {
		return
	}
}
