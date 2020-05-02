package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/smtp"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx"
)

const ValidArgsNum = 2
const LogFile = "./info.log"
const ConfigFile = "./config.json"

type MailConfig struct {
	Email       string `json:"mail-address"`
	Password    string `json:"password"`
	MailTitle   string `json:"mail-title"`
	MailMessage string `json:"mail-message"`
}

type DestMailInfo struct {
	Email      string
	Message    string
	NumOfProxy int
}

func main() {
	logfile, err := os.OpenFile(LogFile, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0666)
	if err != nil {
		failOnError(fmt.Sprintf("%sのオープンに失敗しました", LogFile), err)
	}
	defer logfile.Close()
	log.SetOutput(io.MultiWriter(logfile, os.Stdout))
	log.SetFlags(log.Ldate | log.Ltime)
	log.Println("本ツールはGmailのアカウント設定で[安全性の低いアプリの許可]が必要です。詳細は下記を参照\nhttps://support.google.com/accounts/answer/6010255\n")

	checkArgs()
	config := loadConfig()
	destMailInfoList := getDestMailInfoList(os.Args[1], config.MailMessage)
	//destMailInfoList := getDestMailInfoList("testdata/address-info.xlsx", config.MailMessage)
	postGmails(destMailInfoList, config)
	waitEnter()
}

func postGmails(destMailInfoList []*DestMailInfo, config *MailConfig) {
	for _, destMailInfo := range destMailInfoList {
		postGmail(
			config.Email,
			destMailInfo.Email,
			config.Password,
			config.MailTitle,
			destMailInfo.Message)
	}
}

func loadConfig() *MailConfig {
	configBytes, err := ioutil.ReadFile(ConfigFile)
	if err != nil {
		failOnError("コンフィグファイル読み取りエラー", err)
	}

	config := &MailConfig{}
	if err := json.Unmarshal(configBytes, config); err != nil {
		failOnError("コンフィグファイルJSONパースエラー", err)
	}

	return config
}

func checkArgs() {
	if len(os.Args) != ValidArgsNum {
		exe, err := os.Executable()
		if err != nil {
			failOnError("exeファイル実行パス取得失敗", err)
		}
		exeName := filepath.Base(exe)
		failOnError(
			fmt.Sprintf(
				"%sにExcelファイルをドラッグ&ドロップしてください",
				exeName),
			nil)
	}
}

func postGmail(srcEmail, destEamil, password, title, message string) {
	auth := smtp.PlainAuth(
		"",
		srcEmail,
		password,
		"smtp.gmail.com",
	)

	log.Printf("Sending email to %s. ProxyCount:%d...\n", destEamil, strings.Count(message, "\n")-1)
	err := smtp.SendMail(
		"smtp.gmail.com:587",
		auth,
		srcEmail,
		[]string{destEamil},
		[]byte(
			fmt.Sprintf(
				"To: %s\r\n"+
					"Subject:%s\r\n"+
					"\r\n"+
					"%s",
				destEamil, title, message),
		),
	)

	if err != nil {
		failOnError(fmt.Sprintf("Mail送信エラー。\n宛先:%s \n件名:%s \n本文:%s",
			destEamil, title, message), err)
	}
	log.Printf("Success sending email to %s. ProxyCount %d.\n", destEamil, strings.Count(message, "\n")-1)
}

func failOnError(errMsg string, err error) {
	//errs := errors.WithStack(err)
	log.Println(errMsg)
	if err != nil {
		log.Printf("%s\n", err.Error())
	}
	waitEnter()
	os.Exit(1)
}

func waitEnter() {
	fmt.Println("エンターを押すと処理を終了します。")
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
}

func getProxyList(proxySheet *xlsx.Sheet) []string {
	var proxyList []string
	for i, row := range proxySheet.Rows {
		if i == 0 {
			continue
		}

		proxy := row.Cells[0].String()
		if proxy == "" {
			continue
		}
		proxyList = append(proxyList, proxy)
	}
	return proxyList
}

func getDestMailInfoList(excelFilePath, mailMessageHeader string) []*DestMailInfo {
	excel, err := xlsx.OpenFile(excelFilePath)
	if err != nil {
		failOnError(fmt.Sprintf("%sのオープンに失敗", excelFilePath), err)
	}

	var destMailInfoList []*DestMailInfo
	totalNeedProxy := 0
	addressSheet := excel.Sheets[0]
	for i, row := range addressSheet.Rows {
		if i == 0 {
			continue
		}

		email := row.Cells[0].String()
		if email == "" {
			continue
		}

		numOfNeedProxy, err := row.Cells[1].Int()
		if err != nil {
			failOnError("プロキシ数取得エラー", err)
		}

		destMailInfo := &DestMailInfo{
			Email:      email,
			Message:    mailMessageHeader,
			NumOfProxy: numOfNeedProxy,
		}
		destMailInfoList = append(destMailInfoList, destMailInfo)
		totalNeedProxy += numOfNeedProxy
	}

	proxySheet := excel.Sheets[1]
	proxyList := getProxyList(proxySheet)

	if len(proxyList) < totalNeedProxy {
		failOnError(
			fmt.Sprintf(
				"プロキシの件数が不足しています。要求数:%d 保持数:%d",
				totalNeedProxy, len(proxyList)),
			nil)
	}

	proxyIndex := 0
	for _, destMailInfo := range destMailInfoList {
		for i := 0; i < destMailInfo.NumOfProxy; i++ {
			destMailInfo.Message += fmt.Sprintf("\n%s", proxyList[proxyIndex])
			proxyIndex++
		}
	}

	return destMailInfoList
}
