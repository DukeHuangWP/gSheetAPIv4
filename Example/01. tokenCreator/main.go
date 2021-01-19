package main

import (
	"encoding/json"
	"io/ioutil"
	"log"

	"github.com/DukeHuangWP/gSheetAPIv4"
)

const (
	googleAppCredentials = "./credentials.json"
	googleAccountToken   = "./token.json"
	googleErrorTitle     = "Google Sheet API 預期錯誤"
)

func main() {

	credentials, err := ioutil.ReadFile(googleAppCredentials)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v\n", err)
		//https://developers.google.com/sheets/api/quickstart/go 需要開啟獲取app 使用google api許可
	}

	token, err := ioutil.ReadFile(googleAccountToken)
	if err != nil {
		log.Printf("Unable to read client secret file: %v\n,%v", googleAccountToken, err)
		var getGSService gSheetAPIv4.GoogleSheet
		token, err = getGSService.CreatAccTokenFromWeb(credentials, googleAccountToken, false)
		if err != nil {
			log.Fatalf("Unable to save oauth token: \n %s (%v)\n", credentials, err)
		}

	} else {
		var js map[string]interface{} //空interface可以容許json格式多餘換行，較人性化
		if json.Unmarshal(token, &js) != nil {
			log.Fatalf("INVALID json: %v\n%v", err, token)
		}
	}

	log.Printf("remember to copy: \n%v\n,%v", googleAppCredentials, googleAccountToken)

	return

}
