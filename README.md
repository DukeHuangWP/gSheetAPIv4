# About
Start using google sheet quickly, a package base on Google Sheet API v4.


## Package Install
```
go get github.com/DukeHuangWP/gSheetAPIv4
```

## Before start
You need a google account and API key,

follow this : https://www.youtube.com/watch?v=HwbIxqN4ljY.

And rename ''credentials.json''


#### Set your Google API License
Frist you need set your Google API License,

overwrite your API key :

``./gSheetAPIv4/Example/01. tokenCreator/credentials.json``

and follow the tokenCreator

```
go run ./gSheetAPIv4/Example/01. tokenCreator/main.go
```

After that you'll get a file named ``token.json``.


## How to test it?
Overwrite your API key and token :

``./gSheetAPIv4/Example/02. example_Read&Update/credentials.json``

``./gSheetAPIv4/Example/02. example_Read&Update/token.json``

and modify your code :

``./gSheetAPIv4/Example/02. example_Read&Update/main.go``

```
yourSpreadSheetsID := "1VEImDfFmCAQraxtNSvxp_2IMC1V0axNxZLBdlzfOtqI"
```

now you can run your test :

```
cd ./gSheetAPIv4/Example/02. example_Read&Update
go run main.go
```

Example Code like this:
```go
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
)

func main() {

	//https://docs.google.com/spreadsheets/d/1VEImDfFmCAQraxtNSvxp_2IMC1V0axNxZLBdlzfOtqI/edit#gid=0
	yourSpreadSheetsID := "1VEImDfFmCAQraxtNSvxp_2IMC1V0axNxZLBdlzfOtqI"

	credentials, err := ioutil.ReadFile(googleAppCredentials)
	if err != nil {
		log.Fatalf("Unable to read client secret file: %v\n", err)
	}

	token, err := ioutil.ReadFile(googleAccountToken)
	if err != nil {
		log.Fatalf("INVALID json: %v\n%v", googleAccountToken, err)
	} else {
		var js map[string]interface{} //空interface可以容許json格式多餘換行，較人性化
		if json.Unmarshal(token, &js) != nil {
			log.Fatalf("INVALID json: %v\n%v", err, token)
		}
	}

	//var gs1 gSheetAPIv4.GoogleSheet
	gs1, err := gSheetAPIv4.NewService(credentials, token, yourSpreadSheetsID, false)
	if err != nil {
		log.Fatalf("google sheet connect fail: \n%v", err)
	}

	gid, err := gs1.GetSheetGIDByIndex(0)
	if err != nil {
		log.Fatalf("google sheet get gid fail: \n%v", err)
	}
	log.Printf("got gid : %v\n", gid)

	sheetName, err := gs1.GetSheetNameByGID(gid)
	if err != nil {
		log.Fatalf("google sheet get name fail: \n%v", err)
	}
	log.Printf("got sheetName : %v\n", sheetName)

	var updateValues [][]interface{}
	row := []interface{}{"AAA", "BBB", "CCC", "DDD"}
	updateValues = append(updateValues, row)
	err = gs1.SheetUpdateValue(sheetName+"!A1", updateValues)
	if err != nil {
		log.Printf("update fail : %v\n", err)
	} else {
		log.Printf("update success!\n")
	}

	readValues, err := gs1.SheetReadValue(sheetName + "!A1:C1")
	if err != nil {
		log.Printf("read fail : %v\n", err)
	} else {
		log.Printf("readValues : %v\n", readValues)
	}

}
```
![Demo](https://github.com/DukeHuangWP/gSheetAPIv4/blob/master/Example/HowTo.png?raw=true)




