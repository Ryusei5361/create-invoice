package main

import (
	"context"
	"fmt"
	"github.com/joho/godotenv"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

func GetValuesInSpreadSheet(srv *sheets.Service, spreadsheetID, rg string) (*sheets.Spreadsheet, error) {
	resp, err := srv.Spreadsheets.Get(spreadsheetID).IncludeGridData(true).Ranges(rg).Do()
	if err != nil {
		return nil, err
	}
	return resp, nil
}

func init() {
	fmt.Println("init")
	err := godotenv.Load("env/spreadsheet.env")
	if err != nil {
		log.Fatalf("Error loading .env file: %s", err)
	}
}

// TODO: スプレットシートのタイトルを変更できるようにする。
// TODO: 交通費の情報がいくつあるか指定しなくても持ってこれて、書き込めるようにしたい。
func main() {

	now := time.Now()
	// 請求日
	billDate := fmt.Sprintf("%d/0%d/%d", now.Year(), now.Month(), now.Day())
	// 請求月
	billMonth := fmt.Sprintf("%d/0%d", now.Year(), now.Month())
	// 給料日
	payDate := fmt.Sprintf("%s/15", billMonth)

	fmt.Println(billMonth)

	// コンストラクタ?を作成
	srv, err := sheets.NewService(context.TODO(), option.WithCredentialsFile("credentials/secret.json"))
	if err != nil {
		log.Fatal(err)
	}

	// スプレットシートの ID を読み込む。ID1 が読み込み、ID2 が書き込み
	spreadsheetID1 := os.Getenv("ID1")
	spreadsheetID2 := os.Getenv("ID2")

	// シートとセルを指定、範囲で指定する場合は A1:B6 のようにする
	readRange1 := fmt.Sprintf("大村 %s!C9", billMonth)      // 勤務時間
	readRange2 := fmt.Sprintf("大村 %s!A20:E28", billMonth) // 交通費情報 (多めにセルを指定しておく)

	// 勤務時間を取得
	wrkHr, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange1)
	if err != nil {
		log.Fatal(err)
	}

	// 交通費の情報を取得
	trsptExpnss, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange2)
	if err != nil {
		log.Fatalf("Can't get value: %w", err)
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

	//sprdSht := &sheets.BatchUpdateSpreadsheetRequest{
	//	Properties: &sheets.SpreadsheetProperties{
	//		Title: "New Spreadsheet", // スプレッドシートの名前
	//	},
	//}

	// 更新範囲と更新値の指定
	//valueRange1 := "N4"
	//values1 := [][]interface{}{
	//	{billDate},
	//}
	//valueRange2 := "M15"
	//values2 := [][]interface{}{
	//	{payDate},
	//}
	//valueRange3 := "J18"
	//values3 := [][]interface{}{
	//	{workHours},
	//}
	//valueRange4 := "A20:A28"
	//values4 := [][]interface{}{
	//	{station[0]},
	//	{station[1]},
	//	{station[2]},
	//	{station[3]},
	//	{station[4]},
	//	{station[5]},
	//	{station[6]},
	//	{station[7]},
	//	{station[8]},
	//}
	//valueRange5 := "J20:J28"
	//values5 := [][]interface{}{
	//	{count[0]},
	//	{count[1]},
	//	{count[2]},
	//	{count[3]},
	//	{count[4]},
	//	{count[5]},
	//	{count[6]},
	//	{count[7]},
	//	{count[8]},
	//}
	//valueRange6 := "L20:L28"
	//values6 := [][]interface{}{
	//	{price[0]},
	//	{price[1]},
	//	{price[2]},
	//	{price[3]},
	//	{price[4]},
	//	{price[5]},
	//	{price[6]},
	//	{price[7]},
	//	{price[8]},
	//}

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data: []*sheets.ValueRange{
			{
				Range:          "N4",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{billDate},
				},
			},
			{
				Range:          "M15",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{payDate},
				},
			},
			{
				Range:          "J18",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{workHours},
				},
			},
			{
				Range:          "A20:A28",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{station[0]},
					{station[1]},
					{station[2]},
					{station[3]},
					{station[4]},
					{station[5]},
					{station[6]},
					{station[7]},
					{station[8]},
				},
			},
			{
				Range:          "J20:J28",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{count[0]},
					{count[1]},
					{count[2]},
					{count[3]},
					{count[4]},
					{count[5]},
					{count[6]},
					{count[7]},
					{count[8]},
				},
			},
			{
				Range:          "L20:L28",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{price[0]},
					{price[1]},
					{price[2]},
					{price[3]},
					{price[4]},
					{price[5]},
					{price[6]},
					{price[7]},
					{price[8]},
				},
			},
		},
	}

	_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, rb).Do()
	//_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, sprdSht).Do()
	if err != nil {
		return
	}
}
