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


## Set your Google API License
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

``go run ./gSheetAPIv4/Example/02. example_Read&Update/main.go``




