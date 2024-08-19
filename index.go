package main

import (
	"fmt"
	"github.com/google/uuid"
	"github.com/joho/godotenv"
	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
	"io"
	"log"
	"net/http"
	"os"
	"strconv"
	"strings"
	"time"
)

type PostBody struct {
	values   string
	mailaddr []string
}

func mailHandler(to []string, filePath string) {
	message := gomail.NewMessage()
	message.SetHeader("From", os.Getenv("MAIL_FROM"))
	message.SetHeader("To", to...)
	message.SetHeader("Subject", "Relatório")
	message.SetBody("text/html", "Segue em anexo o relatório solicitado")

	if filePath != "" {
		message.Attach(filePath)
	}

	port, _ := strconv.Atoi(os.Getenv("MAIL_PORT"))
	dialer := gomail.NewDialer(os.Getenv("MAIL_HOST"), port, os.Getenv("MAIL_USER"), os.Getenv("MAIL_PASS"))

	if err := dialer.DialAndSend(message); err != nil {
		log.Fatal(err)
	}
}

func socketHandler(writer http.ResponseWriter, request *http.Request) (filename string, mailaddr []string) {
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
			return "", nil
		}
	}

	return fileName, mailAddr
}

func sheetWriter(jsonStr gjson.Result) (filename string, err error) {
	file := excelize.NewFile()

	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	alfabeto := [26]string{
		"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
	}

	vertical := 2

	gjson.Parse(jsonStr.Str).ForEach(func(key, value gjson.Result) bool {
		horizontal := 0

		gjson.Parse(value.String()).ForEach(func(key, value gjson.Result) bool {
			if vertical == 2 {
				_ = file.SetCellValue("Sheet1", strings.Join([]string{alfabeto[horizontal], strconv.Itoa(1)}, ""), key)
			}

			column := strings.Join([]string{alfabeto[horizontal], strconv.Itoa(vertical)}, "")

			flt, floatErr := strconv.ParseFloat(value.Str, 64)
			date, dateErr := time.Parse("2006-01-02", value.Str)
			datetime, dateTimeErr := time.Parse("2006-01-01 14:14:32", value.Str)

			if floatErr == nil && len(value.String()) < 11 {
				_ = file.SetCellValue("Sheet1", column, flt)
			} else if dateTimeErr != nil {
				_ = file.SetCellValue("Sheet1", column, datetime.Format("02/01/2006 14:14:32"))
			} else if dateErr == nil {
				_ = file.SetCellValue("Sheet1", column, date.Format("02/01/2006"))
			} else if value.Type == gjson.String {
				_ = file.SetCellValue("Sheet1", column, value.Str)
			} else if value.Type == gjson.Number {
				_ = file.SetCellValue("Sheet1", column, value.Num)
			} else {
				_ = file.SetCellValue("Sheet1", column, value)
			}

			horizontal++
			return true
		})

		vertical++
		return true
	})

	uuidInfo, _ := uuid.NewUUID()
	fileName := uuidInfo.String() + ".xlsx"

	err = file.SaveAs(fileName)
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

	http.HandleFunc("/", func(writer http.ResponseWriter, request *http.Request) {
		filename, mailAddr := socketHandler(writer, request)
		if filename != "" {
			mailHandler(mailAddr, filename)
		}
	})

	httpErr := http.ListenAndServe(":8081", nil)
	if httpErr != nil {
		log.Fatal("Erro ao iniciar o servidor na porta 8081")
		return
	}
}
