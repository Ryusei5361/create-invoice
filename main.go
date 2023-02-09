package main

import (
	"context"
	"fmt"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"os"
	"time"
)

var spreadsheetID1 = "1SAwRc11TMl9fc8243Es0HvL8ZK3SHa6nEyTprCBC6Bk"
var spreadsheetID2 = "1jo4DVvChI5sNXFlyUhqF8oSmfTVyiryhiJy6KHkd_xw"

type SheetClient struct {
	srv           *sheets.Service
	spreadsheetID string
}

func NewSheetClient(spreadsheetID string) (*SheetClient, error) {
	credential := option.WithCredentialsFile("credentials/secret.json")
	srv, err := sheets.NewService(context.TODO(), credential)
	if err != nil {
		log.Fatal(err)
	}

	return &SheetClient{
		srv:           srv,
		spreadsheetID: spreadsheetID,
	}, nil
}

func (s *SheetClient) Get(range_ string) (*sheets.Spreadsheet, error) {
	resp, err := s.srv.Spreadsheets.Get(s.spreadsheetID).IncludeGridData(true).Ranges(range_).Do()
	if err != nil {
		return nil, err
	}
	return resp, nil
}

func main() {

	now := time.Now()
	date := fmt.Sprintf("%d/%d/%d", now.Year(), int(now.Month()), now.Day())
	transferredDate := fmt.Sprintf("%d/%d/15", now.Year(), int(now.Month()))

	//fmt.Println(date)

	client, err := NewSheetClient(os.Getenv(spreadsheetID1))
	if err != nil {
		log.Fatal(err)
	}
	//credential := option.WithCredentialsFile("credentials/secret.json")
	//srv, err := sheets.NewService(context.TODO(), credential)
	//if err != nil {
	//	log.Fatal(err)
	//}

	//spreadSheet := &sheets.Spreadsheet{
	//	Properties: &sheets.SpreadsheetProperties{
	//		Title:    "New Spreadsheet", // スプレッドシートの名前
	//		Locale:   "ja_JP",           // ロケール
	//		TimeZone: "Asia/Tokyo",      // タイムゾーン
	//	},
	//}

	// スプレッドシートを新規作成
	//createResponse, err := srv.Spreadsheets.Create(spreadSheet).Do()
	//if err != nil {
	//	log.Fatal(err)
	//}

	// シートとセルを指定、範囲で指定する場合は A1:B6 のようにする
	readRange1 := "大村 2023/02!C9"
	readRange2 := "大村 2023/02!A21:E23"

	rg := &sheets.BatchGetValuesResponse{
		ValueRanges: []*sheets.ValueRange{
			{
				Range:          readRange1,
				MajorDimension: "ROWS",
			},
			{
				Range:          readRange2,
				MajorDimension: "ROWS",
			},
		},
	}

	// 値を取得
	value, err := client.Get(readRange1)
	//resp, err := srv.Spreadsheets.Get(spreadsheetID1).IncludeGridData(true).Ranges(rg...).Do()
	if err != nil {
		log.Fatal(err)
	}
	// interface 型のフォーマット演算子 %#v
	//fmt.Printf("%#v\n", resp)

	//station = [];
	// セル情報の取得
	for _, s := range resp.Sheets {
		for _, row := range s.Data[0].RowData {
			for _, value := range row.Values {
				//fmt.Printf("%v %v %v %v %v\n", row.Values[0].FormattedValue, row.Values[1].FormattedValue,
				//	row.Values[2].FormattedValue, row.Values[3].FormattedValue, row.Values[4].FormattedValue)
				//fmt.Println(value.EffectiveValue.StringValue)
				//fmt.Println(value.UserEnteredValue.StringValue)
				fmt.Println(value.FormattedValue)
			}
		}
	}

	// スプレットシートから読み込んだセルの値、今回だと勤務時間
	//var workingHours = resp.Sheets[0].Data[0].RowData[0].Values[0].FormattedValue
	//// 25:50 のように、: が含まれる文字列を小数に変換するために : から . へ置き換える
	//workingHours = strings.Replace(workingHours, ":", ".", 1)
	//// 小数として扱いたいので、string 型を float64 型に変換
	//newWorkingHours, err := strconv.ParseFloat(workingHours, 64)
	//if err != nil {
	//	log.Fatalln(err)
	//}

	//fmt.Println(newWorkingHours)

	// 請求書の勤務時間を書き込むセルを指定
	//writeRange1 := "請求書!J18"
	//writeRange2 := "請求書!M15"
	// 請求書の勤務時間載せるに書き込む内容
	//vr := &sheets.ValueRange{
	//	Values: [][]interface{}{
	//		{newWorkingHours},
	//	},
	//}

	// 更新範囲と更新値の指定
	//valueRange1 := "J18"
	//values1 := [][]interface{}{
	//	{newWorkingHours},
	//}

	valueRange2 := "N4"
	values2 := [][]interface{}{
		{date},
	}

	valueRange3 := "M15"
	values3 := [][]interface{}{
		{transferredDate},
	}

	//valueRange4 := "A20:J22"
	//values4 := [][]interface{}{
	//	{transferredDate}, {transferredDate},
	//}

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data: []*sheets.ValueRange{
			//{
			//	Range:          valueRange1,
			//	MajorDimension: "ROWS",
			//	Values:         values1,
			//},
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
			//{
			//	Range:          valueRange4,
			//	MajorDimension: "ROWS",
			//	Values:         values4,
			//},
		},
	}

	_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, rb).Do()
	if err != nil {
		return
	}

	//fmt.Println(vr)

	// 書き込む
	//_, err = srv.Spreadsheets.Values.Update(spreadsheetID2, writeRange1, vr).ValueInputOption("RAW").Do()
	//if err != nil {
	//	log.Fatalln(err)
	//}
	//_, err = srv.Spreadsheets.Values.Update(spreadsheetID2, writeRange2, date).ValueInputOption("RAW").Do()
	//if err != nil {
	//	log.Fatalln(err)
	//}
}
