package main

import (
	"bufio"
	"context"
	"fmt"
	"github.com/joho/godotenv"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
	"time"
)

//const name = "阿部"

func GetValuesInSpreadSheet(srv *sheets.Service, spreadsheetID, rg string) (*sheets.Spreadsheet, error) {
	resp, err := srv.Spreadsheets.Get(spreadsheetID).IncludeGridData(true).Ranges(rg).Do()
	if err != nil {
		return nil, err
	}
	return resp, nil
}

func init() {
	err := godotenv.Load("env/spreadsheet.env")
	if err != nil {
		log.Fatalf("Error loading .env file: %s", err)
	}
}

// TODO: スプレットシートのタイトルを変更できるようにする。
// TODO: 交通費の情報がいくつあるか指定しなくても持ってこれて、書き込めるようにしたい。
func main() {
	// 入力を受け付ける
	sc := bufio.NewScanner(os.Stdin)
	fmt.Print("名前を入力してください\n>>> ")
	sc.Scan()
	// 誰の請求書か
	name := sc.Text()

	now := time.Now()
	// 請求日
	billDate := fmt.Sprintf("%d/0%d/%d", now.Year(), now.Month(), now.Day())
	// 請求月
	billMonth := fmt.Sprintf("%d/0%d", now.Year(), now.Month()-1)
	// 給料日
	payDate := fmt.Sprintf("%d/%d/15", now.Year(), now.Month())

	// コンストラクタ?を作成
	srv, err := sheets.NewService(context.TODO(), option.WithCredentialsFile("credentials/secret.json"))
	if err != nil {
		log.Fatal(err)
	}

	// スプレットシートの ID を読み込む。ID1 が読み込み、ID2 が書き込み
	spreadsheetID1 := os.Getenv("ID1")
	spreadsheetID2 := os.Getenv("ID2")

	// シートとセルを指定、範囲で指定する場合は A1:B6 のようにする
	readRange1 := fmt.Sprintf("%s %s!C9", name, billMonth)      // 勤務時間
	readRange2 := fmt.Sprintf("%s %s!A20:E28", name, billMonth) // 交通費情報 (多めにセルを指定しておく)

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

	// 交通費情報の長さ
	trafficlenght := len(trsptExpnss.Sheets[0].Data[0].RowData)

	// 駅名、回数、値段の変数
	station := make([][]interface{}, trafficlenght, trafficlenght)
	count := make([][]interface{}, trafficlenght, trafficlenght)
	price := make([][]interface{}, trafficlenght, trafficlenght)

	// 駅名、回数、値段を変数に格納
	for _, s := range trsptExpnss.Sheets {
		for i, row := range s.Data[0].RowData {
			station[i] = append(station[i], fmt.Sprintf("%v %v %v", row.Values[0].FormattedValue, row.Values[1].FormattedValue,
				row.Values[2].FormattedValue))
			count[i] = append(count[i], row.Values[3].FormattedValue)
			price[i] = append(price[i], row.Values[4].FormattedValue)
		}
	}

	// スプレットシートに書き込むための範囲と値
	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data: []*sheets.ValueRange{
			{
				Range:          name + "!N4",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{billDate},
				},
			},
			{
				Range:          name + "!M15",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{"お支払い期限: " + payDate},
				},
			},
			{
				Range:          name + "!J18",
				MajorDimension: "ROWS",
				Values: [][]interface{}{
					{math.Ceil(workHours)},
				},
			},
			{
				Range:          name + "!A20:A28",
				MajorDimension: "ROWS",
				Values:         station,
			},
			{
				Range:          name + "!J20:J28",
				MajorDimension: "ROWS",
				Values:         count,
			},
			{
				Range:          name + "!L20:L28",
				MajorDimension: "ROWS",
				Values:         price,
			},
		},
	}

	// スプレットシートに値を書き込む
	_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, rb).Do()
	if err != nil {
		log.Fatalln(err)
	}
}
