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

func main() {

	now := time.Now()
	// 請求日
	billdate := fmt.Sprintf("%d/%d/%d", now.Year(), int(now.Month()), now.Day())
	// 給料日
	payDate := fmt.Sprintf("%d/%d/15", now.Year(), int(now.Month()))

	//fmt.Println(date)

	srv, err := sheets.NewService(context.TODO(), option.WithCredentialsFile("credentials/secret.json"))
	if err != nil {
		log.Fatal(err)
	}

	// シートとセルを指定、範囲で指定する場合は A1:B6 のようにする
	readRange1 := "大村 2023/02!C9"
	readRange2 := "大村 2023/02!A21:E23"

	// 値を取得
	wrkHr, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange1)
	if err != nil {
		log.Fatal(err)
	}

	trsptExpnss, err := GetValuesInSpreadSheet(srv, spreadsheetID1, readRange2)
	if err != nil {
		log.Fatal(err)
	}

	// スプレットシートから読み込んだセルの値、今回だと勤務時間
	var workingHours = wrkHr.Sheets[0].Data[0].RowData[0].Values[0].FormattedValue
	// 25:50 のように、: が含まれる文字列を小数に変換するために : から . へ置き換える
	workingHours = strings.Replace(workingHours, ":", ".", 1)
	// 小数として扱いたいので、string 型を float64 型に変換
	newWorkingHours, err := strconv.ParseFloat(workingHours, 64)
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Println(newWorkingHours)

	for _, s := range trsptExpnss.Sheets {
		for _, row := range s.Data[0].RowData {
			fmt.Printf("%v %v %v %v %v\n", row.Values[0].FormattedValue, row.Values[1].FormattedValue,
				row.Values[2].FormattedValue, row.Values[3].FormattedValue, row.Values[4].FormattedValue)
		}
	}

	// 更新範囲と更新値の指定
	valueRange2 := "N4"
	values2 := [][]interface{}{
		{billdate},
	}

	valueRange3 := "M15"
	values3 := [][]interface{}{
		{payDate},
	}

	rb := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data: []*sheets.ValueRange{
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
		},
	}

	_, err = srv.Spreadsheets.Values.BatchUpdate(spreadsheetID2, rb).Do()
	if err != nil {
		return
	}
}
