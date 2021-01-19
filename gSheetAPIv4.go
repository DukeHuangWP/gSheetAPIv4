package gSheetAPIv4

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"os"

	"golang.org/x/net/context"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/sheets/v4"
)

const (
	googleErrorTitle = "Google Sheet API 預期錯誤"
)

type GoogleSheet struct {
	Srv            *sheets.Service
	Ctx            context.Context
	AppCredentials []byte //程式本身許可license
	AccToken       []byte //google帳戶license
	SpreadSheetID  string //預設google Sheet ID (由網址url取得)
}

//獲取google API 控制權限設定(綁定method)
func (gs *GoogleSheet) getGoogleAPIConfig(appCredentials []byte, setReadOnly bool) (config *oauth2.Config, err error) {
	// https://developers.google.com/sheets/api/guides/authorizing 選擇api控制google帳號權限
	// https://www.googleapis.com/auth/spreadsheets.readonly	Allows read-only access to the user's sheets and their properties.
	// https://www.googleapis.com/auth/spreadsheets	Allows read/write access to the user's sheets and their properties.
	// https://www.googleapis.com/auth/drive.readonly	Allows read-only access to the user's file metadata and file content.
	// https://www.googleapis.com/auth/drive.file	Per-file access to files created or opened by the app.
	// https://www.googleapis.com/auth/drive
	var scopeUrl string
	if setReadOnly {
		scopeUrl = sheets.SpreadsheetsReadonlyScope
	} else {
		scopeUrl = sheets.SpreadsheetsScope
	}

	config, err = google.ConfigFromJSON(appCredentials, scopeUrl)
	if err != nil {
		//log.Printf("Unable to parse client secret file to config: %v\n", err)
		return nil, errors.New(googleErrorTitle + ": AppCredentials 內容格式錯誤 !")
	}

	return config, nil
}

//轉劇矩陣，Google習慣使用y,x表示平面，需轉致成x,y
func transposeMatrix(slice [][]interface{}) [][]interface{} {
	xl := len(slice[0])
	yl := len(slice)
	result := make([][]interface{}, xl)
	for i := range result {
		result[i] = make([]interface{}, yl)
	}
	for i := 0; i < xl; i++ {
		for j := 0; j < yl; j++ {
			result[i][j] = slice[j][i]
		}
	}
	return result
}

func NewService(appCredentials []byte, accToken []byte, spreadSheetID string, setReadOnly bool) (GoogleSheet, error) {

	var gSheet GoogleSheet
	gSheet.AppCredentials = appCredentials
	gSheet.AccToken = accToken
	gSheet.SpreadSheetID = spreadSheetID //預先存入docs網址方便日後存取

	config, err := gSheet.getGoogleAPIConfig(appCredentials, setReadOnly)
	if err != nil {
		// 	log.Printf("Unable to parse client secret file to config: %v\n", err)
		return GoogleSheet{}, errors.New(googleErrorTitle + ": AppCredentials 內容格式錯誤 ! >> " + err.Error())
	}

	tok := &oauth2.Token{}
	err = json.NewDecoder(bytes.NewReader(gSheet.AccToken)).Decode(tok)
	if err != nil {
		//log.Fatalf("Unable to cache oauth token: %v", err)
		return GoogleSheet{}, errors.New(googleErrorTitle + ": AccountToken 內容格式錯誤 ! >> " + err.Error())
	}
	client := config.Client(context.Background(), tok)

	gSheet.Srv, err = sheets.New(client)
	if err != nil {
		//log.Printf("Unable to retrieve Sheets client: %v\n", err)
		return GoogleSheet{}, err
	}

	return gSheet, nil

}

//設置Google Sheet連結服務
func (gs *GoogleSheet) SetService(spreadSheetID string, setReadOnly bool) (*sheets.Service, error) {

	gs.SpreadSheetID = spreadSheetID //預先存入docs網址方便日後存取

	config, err := gs.getGoogleAPIConfig(gs.AppCredentials, setReadOnly)
	if err != nil {
		// 	log.Printf("Unable to parse client secret file to config: %v\n", err)
		return nil, errors.New(googleErrorTitle + ": AppCredentials 內容格式錯誤 ! >> " + err.Error())
	}

	tok := &oauth2.Token{}
	err = json.NewDecoder(bytes.NewReader(gs.AccToken)).Decode(tok)
	if err != nil {
		//log.Fatalf("Unable to cache oauth token: %v", err)
		return nil, errors.New(googleErrorTitle + ": AccountToken 內容格式錯誤 ! >> " + err.Error())
	}
	client := config.Client(context.Background(), tok)

	gs.Srv, err = sheets.New(client)
	if err != nil {
		//log.Printf("Unable to retrieve Sheets client: %v\n", err)
		return nil, err
	}

	return gs.Srv, nil

}

//由Sheet表名稱找出Sheet GID (使用名稱)
func (gs *GoogleSheet) GetSheetGIDByName(sheetName string) (int64, error) {

	sheetService, err := gs.Srv.Spreadsheets.Get(gs.SpreadSheetID).Do()
	if err != nil {
		return 0, errors.New(googleErrorTitle + ": 「" + gs.SpreadSheetID + "」 該SheetID不存在，請檢查網址或AccountToken權限!")
	}

	for _, sheet := range sheetService.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}
	return 0, errors.New(googleErrorTitle + ": 「" + sheetName + "」 該Sheet Name不存在，請檢查Sheet內容!")
}

//由Sheet表名稱找出Sheet GID (使用index)
func (gs *GoogleSheet) GetSheetGIDByIndex(sheetIndex int64) (int64, error) {

	sheetService, err := gs.Srv.Spreadsheets.Get(gs.SpreadSheetID).Do()
	if err != nil {
		return 0, errors.New(googleErrorTitle + ": 「" + gs.SpreadSheetID + "」 該SheetID不存在，請檢查網址或AccountToken權限!")
	}

	for _, sheet := range sheetService.Sheets {
		if sheet.Properties.Index == sheetIndex {
			return sheet.Properties.SheetId, nil
		}
	}
	return 0, errors.New(googleErrorTitle + ": 「" + fmt.Sprint(sheetIndex) + "」 該Sheet Index不存在，請檢查Sheet內容!")
}

//由Sheet表index找出Sheet名稱 (使用index)
func (gs *GoogleSheet) GetSheetNameByIndex(sheetIndex int64) (string, error) {

	sheetService, err := gs.Srv.Spreadsheets.Get(gs.SpreadSheetID).Do()
	if err != nil {
		return "", errors.New(googleErrorTitle + ": 「" + gs.SpreadSheetID + "」 該SheetID不存在，請檢查網址或AccountToken權限!")
	}

	for _, sheet := range sheetService.Sheets {
		if sheet.Properties.Index == sheetIndex {
			return sheet.Properties.Title, nil
		}
	}
	return "", errors.New(googleErrorTitle + ": 「" + fmt.Sprint(sheetIndex) + "」 該Sheet Index不存在，請檢查Sheet內容!")
}

//由Sheet表GID找出Sheet名稱 (使用GID)
func (gs *GoogleSheet) GetSheetNameByGID(sheetGID int64) (string, error) {

	sheetService, err := gs.Srv.Spreadsheets.Get(gs.SpreadSheetID).Do()
	if err != nil {
		return "", errors.New(googleErrorTitle + ": 「" + gs.SpreadSheetID + "」 該SheetID不存在，請檢查網址或AccountToken權限!")
	}

	for _, sheet := range sheetService.Sheets {
		if sheet.Properties.SheetId == sheetGID {
			return sheet.Properties.Title, nil
		}
	}
	return "", errors.New(googleErrorTitle + ": 「" + fmt.Sprint(sheetGID) + "」 該SheetGID不存在，請檢查Sheet內容!")
}

//建立Google帳戶Token File(需要瀏覽器手動操作)
func (gs *GoogleSheet) CreatAccTokenFromWeb(appCredentials []byte, accTokenFilePath string, setReadOnly bool) (accToken []byte, err error) {

	config, err := gs.getGoogleAPIConfig(appCredentials, setReadOnly)
	if err != nil {
		//log.Printf("Unable to parse client secret file to config: %v\n", err)
		//https://console.developers.google.com/apis/credentials
		return nil, errors.New(googleErrorTitle + ": AppCredentials 內容格式錯誤 !")
	}

	// Request a token from the web, then returns the retrieved token.
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("\n%v\n↑請使用瀏覽器設定你的Google帳戶，並貼入你的token authorization code ↑\n", authURL)
	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		//log.Fatalf("Unable to read authorization code: %v", err)
		return nil, err
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		//log.Fatalf("Unable to retrieve token from web: %v", err)
		return nil, err
	}

	// Saves a token to a file path.
	fmt.Printf("已將token authorization code存入 : %s\n", accTokenFilePath)
	accTokenFileWrite, err := os.OpenFile(accTokenFilePath, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		//log.Fatalf("Unable to cache oauth token: %v", err)
		return nil, errors.New(googleErrorTitle + ": 「" + accTokenFilePath + "」AccountToken儲存失敗!")
	}
	defer accTokenFileWrite.Close()
	json.NewEncoder(accTokenFileWrite).Encode(tok)

	gs.Ctx = context.Background()
	client := config.Client(gs.Ctx, tok)
	_, err = sheets.New(client)
	if err != nil {
		//log.Fatalf("Unable to retrieve Sheets client: %v", err)
		return nil, err
	}

	token, err := ioutil.ReadFile(accTokenFilePath)
	if err != nil {
		//log.Fatalf("Unable to save oauth token: %v", err)
		return nil, err
	}

	return []byte(token), nil
}

//獲取sheet Range中的數值
func (gs *GoogleSheet) SheetReadValue(readRange string) ([][]interface{}, error) {
	resp, err := gs.Srv.Spreadsheets.Values.Get(gs.SpreadSheetID, readRange).Do()
	if err != nil {
		return nil, err
	}

	return resp.Values, err
}

//獲取sheet Range中的公式
func (gs *GoogleSheet) SheetReadFormula(readRange string) ([][]interface{}, error) {
	resp, err := gs.Srv.Spreadsheets.Values.Get(gs.SpreadSheetID, readRange).ValueRenderOption("FORMULA").Do()
	if err != nil {
		return nil, err
	}

	return resp.Values, err
}

//寫入sheet Range中的數值
func (gs *GoogleSheet) SheetWriteValue(writeRange string, values [][]interface{}) error {
	valueInputOption := "RAW"
	requestBody := &sheets.ValueRange{
		MajorDimension: "ROWS",
		Values:         values,
	}

	_, err := gs.Srv.Spreadsheets.Values.Append(gs.SpreadSheetID, writeRange, requestBody).ValueInputOption(valueInputOption).Do()

	return err
}

//寫入sheet Range中的公式
func (gs *GoogleSheet) SheetWriteFormula(writeRange string, values [][]interface{}) error {
	valueInputOption := "USER_ENTERED"
	requestBody := &sheets.ValueRange{
		MajorDimension: "ROWS",
		Values:         values,
	}

	_, err := gs.Srv.Spreadsheets.Values.Append(gs.SpreadSheetID, writeRange, requestBody).ValueInputOption(valueInputOption).Do()

	return err
}

//改變sheet Range中的數值
func (gs *GoogleSheet) SheetUpdateValue(updateRange string, updateValues [][]interface{}) error {
	valueInputOption := "RAW"
	requestBody := &sheets.ValueRange{
		MajorDimension: "ROWS",
		Values:         updateValues,
	}

	_, err := gs.Srv.Spreadsheets.Values.Update(gs.SpreadSheetID, updateRange, requestBody).ValueInputOption(valueInputOption).Do()

	return err
}

//改變sheet Range中的公式
func (gs *GoogleSheet) SheetUpdateFormula(updateRange string, updateValues [][]interface{}) error {
	valueInputOption := "USER_ENTERED"
	requestBody := &sheets.ValueRange{
		MajorDimension: "ROWS",
		Values:         updateValues,
	}

	_, err := gs.Srv.Spreadsheets.Values.Update(gs.SpreadSheetID, updateRange, requestBody).ValueInputOption(valueInputOption).Do()

	return err
}

//清除sheet Range中的數值
func (gs *GoogleSheet) SheetClear(clearRange string) error {
	// rb has type *ClearValuesRequest
	requestBody := &sheets.ClearValuesRequest{}

	_, err := gs.Srv.Spreadsheets.Values.Clear(gs.SpreadSheetID, clearRange, requestBody).Do()

	return err
}

//重新命名 sheet 名稱 (由GID)
func (gs *GoogleSheet) SheetRenameByGID(sheetGID int64, newName string) (resp *sheets.BatchUpdateSpreadsheetResponse, err error) {

	requestBody := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&sheets.Request{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Properties: &sheets.SheetProperties{
					SheetId: sheetGID,
					Title:   newName,
				},
				Fields: "title",
			},
		}},
	}

	resp, err = gs.Srv.Spreadsheets.BatchUpdate(gs.SpreadSheetID, requestBody).Context(gs.Ctx).Do()
	return resp, err
}

//隱藏或顯示 sheet (由GID)
func (gs *GoogleSheet) SheetHideByGID(sheetGID int64, hideSheet bool) (resp *sheets.BatchUpdateSpreadsheetResponse, err error) {

	requestBody := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&sheets.Request{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Properties: &sheets.SheetProperties{
					SheetId: sheetGID,
					Hidden:  hideSheet,
				},
				Fields: "hidden",
			},
		}},
	}

	resp, err = gs.Srv.Spreadsheets.BatchUpdate(gs.SpreadSheetID, requestBody).Context(gs.Ctx).Do()
	return resp, err
}

//複製後並貼上sheet Range中的數值或公式
func (gs *GoogleSheet) SheetCopyPasteByGID(sourceGID, destGID int64, sourceStartPos, sourceEndPos, destStartPos, destEndPos [2]int64, copyWithFormat bool) (resp *sheets.BatchUpdateSpreadsheetResponse, err error) {
	var pasteType string
	if copyWithFormat {
		pasteType = "PASTE_NORMAL"
	} else {
		pasteType = "PASTE_FORMULA"
	} //https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#PasteType

	requestBody := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&sheets.Request{
			CopyPaste: &sheets.CopyPasteRequest{
				Source: &sheets.GridRange{
					SheetId:          sourceGID,
					StartColumnIndex: sourceStartPos[1] - 1,
					EndColumnIndex:   sourceEndPos[1],
					StartRowIndex:    sourceStartPos[0] - 1,
					EndRowIndex:      sourceEndPos[0],
				},
				Destination: &sheets.GridRange{
					SheetId:          destGID,
					StartColumnIndex: destStartPos[1] - 1,
					EndColumnIndex:   destEndPos[1],
					StartRowIndex:    destStartPos[0] - 1,
					EndRowIndex:      destEndPos[0],
				},
				PasteType:        pasteType,
				PasteOrientation: "NORMAL",
			},
		}},
	}

	resp, err = gs.Srv.Spreadsheets.BatchUpdate(gs.SpreadSheetID, requestBody).Context(gs.Ctx).Do()
	return resp, err
}

//操作兩Google Sheet 間作sheet複製(需要自行設置權限,並取得Sheet的gid)
func (gs *GoogleSheet) CopyBetweenSheet(sourceSheetID string, sourceSheetGID int64, destSheetID string) (resp *sheets.SheetProperties, err error) {

	requestBody := &sheets.CopySheetToAnotherSpreadsheetRequest{
		DestinationSpreadsheetId: destSheetID,
	}

	resp, err = gs.Srv.Spreadsheets.Sheets.CopyTo(sourceSheetID, sourceSheetGID, requestBody).Context(gs.Ctx).Do()

	return resp, err
}

//複製外部sheet到現在工作的Google Sheet(需要自行設置權限,並取得Sheet的gid)
func (gs *GoogleSheet) CopyFromSheet(sourceSheetID string, sourceSheetGID int64, newName string) (resp *sheets.SheetProperties, err error) {

	requestBody := &sheets.CopySheetToAnotherSpreadsheetRequest{
		DestinationSpreadsheetId: gs.SpreadSheetID,
	}

	resp, err = gs.Srv.Spreadsheets.Sheets.CopyTo(sourceSheetID, sourceSheetGID, requestBody).Context(gs.Ctx).Do()
	if newName != "" && err == nil {
		_, err = gs.SheetRenameByGID(resp.SheetId, newName)
		if err == nil {
			_, err = gs.SheetHideByGID(resp.SheetId, false)
		}
	}

	return resp, err
}
