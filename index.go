package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/google/uuid"
	"github.com/joho/godotenv"
	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
)

type Item struct {
	file     string
	mailaddr []string
}

func mailHandler(params *Item) {
	message := gomail.NewMessage()
	message.SetHeader("From", os.Getenv("MAIL_FROM_ADDRESS"))
	message.SetHeader("To", params.mailaddr...)
	message.SetHeader("Subject", "Relatório")
	message.SetBody("text/html", "Segue em anexo o relatório solicitado")

	if params.file != "" {
		message.Attach(params.file)
	}

	port, _ := strconv.Atoi(os.Getenv("MAIL_PORT"))
	dialer := gomail.NewDialer(os.Getenv("MAIL_HOST"), port, os.Getenv("MAIL_USERNAME"), os.Getenv("MAIL_PASSWORD"))

	if err := dialer.DialAndSend(message); err != nil {
		log.Fatal(err)
	}
}

func socketHandler(writer http.ResponseWriter, request *http.Request) *Item {
	buf := new(strings.Builder)
	_, err := io.Copy(buf, request.Body)
	if err != nil {
		fmt.Println(err)
	}

	fileName := ""
	var mailAddr []string
	var fileNameErr error = nil

	gjson.Parse(buf.String()).ForEach(func(key, value gjson.Result) bool {
		if key.Str == "values" {
			fileName, fileNameErr = sheetWriter(value)
		} else if key.Str == "mailaddr" {
			gjson.Parse(value.String()).ForEach(func(key, value gjson.Result) bool {
				mailAddr = append(mailAddr, value.String())
				return true
			})
		}

		return true
	})

	if fileNameErr != nil {
		_, err := writer.Write([]byte(fileNameErr.Error()))
		if err != nil {
			return &Item{"", nil}
		}
	}

	return &Item{fileName, mailAddr}
}

func sheetWriter(jsonStr gjson.Result) (filename string, err error) {
	sheet := excelize.NewFile()

	defer func() {
		if err := sheet.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	alfabeto := [26]string{
		"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	}

	vertical := 2

	gjson.Parse(jsonStr.String()).ForEach(func(key, value gjson.Result) bool {
		horizontal, horizontalIdx2 := 0, 0

		gjson.Parse(value.String()).ForEach(func(key, value gjson.Result) bool {
			column := ""
			if horizontal < 26 {
				column = fmt.Sprintf("%s%d", alfabeto[horizontal], vertical)
			} else {
				column = fmt.Sprintf("%s%s%d", alfabeto[horizontalIdx2], alfabeto[horizontal%26], vertical)
			}

			if vertical == 2 {
				regex := regexp.MustCompile("[0-9]+")
				regexSearch := regex.ReplaceAllString(column, "")
				header := fmt.Sprintf("%s%d", regexSearch, 1)
				_ = sheet.SetCellValue("Sheet1", header, key.String())
			}

			flt, floatErr := strconv.ParseFloat(value.Str, 64)
			date, dateErr := time.Parse("2006-01-02", value.Str)
			datetime, dateTimeErr := time.Parse("2006-01-01 14:14:32", value.Str)

			if floatErr == nil && len(value.String()) < 11 {
				_ = sheet.SetCellValue("Sheet1", column, flt)
			} else if dateTimeErr == nil {
				_ = sheet.SetCellValue("Sheet1", column, datetime.Format("02/01/2006 14:14:32"))
			} else if dateErr == nil {
				_ = sheet.SetCellValue("Sheet1", column, date.Format("02/01/2006"))
			} else if value.Type == gjson.String {
				_ = sheet.SetCellValue("Sheet1", column, value.Str)
			} else if value.Type == gjson.Number {
				_ = sheet.SetCellValue("Sheet1", column, value.Num)
			} else {
				_ = sheet.SetCellValue("Sheet1", column, value)
			}

			horizontal++
			if horizontal%26 == 0 && horizontal >= 52 {
				horizontalIdx2++
			}

			return true
		})

		vertical++
		return true
	})

	uuidInfo, _ := uuid.NewUUID()
	fileName := uuidInfo.String() + ".xlsx"

	err = sheet.SaveAs(fileName)
	if err != nil {
		return "", err
	}

	return fileName, nil
}

func main() {
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Erro ao carregar o arquivo .env")
	}

	fmt.Println("Iniciando servidor na porta 8081")

	go http.HandleFunc("/", func(writer http.ResponseWriter, request *http.Request) {
		fmt.Println("Conexão recebida: ", request.RemoteAddr)
		item := socketHandler(writer, request)
		if item.mailaddr != nil {
			mailHandler(item)
		}

		if item.file != "" {
			err := os.Remove(item.file)
			if err != nil {
				log.Fatal(err)
			}
		}
	})

	httpErr := http.ListenAndServe(":8081", nil)
	if httpErr != nil {
		log.Fatal("Erro ao iniciar o servidor na porta 8081")
		return
	}
}
