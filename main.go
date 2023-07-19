package main

import (
	"encoding/xml"
	"fmt"
	"github.com/spf13/cast"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"regexp"
)

type MyResponse struct {
	Message string `xml:"message"`
}

func main() {
	fileName := "5starStocks.txt"
	file, err := os.OpenFile(fileName, os.O_WRONLY|os.O_CREATE|os.O_APPEND, 0644)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	for _, stock := range openExcel() {
		parseStock(file, stock)
	}
}

func openExcel() []string {
	filePath := "nse.xlsx"

	// Open the XLSX file
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Error opening XLSX file: %v", err)
	}

	// Specify the sheet index from which you want to read the column
	sheetIndex := 0

	// Specify the column index (0-based) from which you want to read the data
	columnIndex := 1 // For example, 1 for the second column

	// Read the column data
	return readColumn(xlFile, sheetIndex, columnIndex)
}

func readColumn(file *xlsx.File, sheetIndex, columnIndex int) []string {
	var columnData []string

	// Get the sheet at the specified index
	sheet := file.Sheets[sheetIndex]

	// Iterate through the rows and read the data from the specified column
	for _, row := range sheet.Rows {
		if columnIndex < len(row.Cells) {
			columnData = append(columnData, row.Cells[columnIndex].String())
		}
	}

	return columnData
}

func parseStock(file *os.File, stock string) {
	url := "https://ticker.finology.in/company/" + stock // Replace with the desired URL

	// Send GET request
	response, err := http.Get(url)
	if err != nil {
		fmt.Printf("Error making the GET request: %s\n", err.Error())
		return
	}
	defer response.Body.Close()

	// Read the response body
	body, err := ioutil.ReadAll(response.Body)
	if err != nil {
		fmt.Printf("Error reading response body: %s\n", err.Error())
		return
	}

	// Parse the response based on the content type
	contentType := response.Header.Get("Content-Type")
	if contentType == "text/html; charset=utf-8" {
		re := regexp.MustCompile(`Valuation Rating is (\d+) out of (\d+)\.`)

		match := re.FindStringSubmatch(string(body))
		var ratingNumber string
		if len(match) == 3 {
			// Extract the rating number from the captured match
			ratingNumber = match[1]
		} else {
			log.Println("Rating number not found for ", stock)
			return
		}

		ratingInt := cast.ToInt(ratingNumber)
		if ratingInt >= 4 {
			_, err = file.WriteString("Stock : " + stock + " rating : " + ratingNumber + "\n")
			if err != nil {
				log.Fatal(err)
			}
		}
	} else if contentType == "application/xml" {
		var parsedResponse MyResponse
		err = xml.Unmarshal(body, &parsedResponse)
		if err != nil {
			fmt.Printf("Error parsing XML response: %s\n", err.Error())
			return
		}
		fmt.Printf("Parsed XML response: %s\n", parsedResponse.Message)
	} else if contentType == "application/json" {
		// Parse JSON response here using the encoding/json package
		// Replace this placeholder code with your actual JSON parsing logic
		fmt.Printf("JSON response: %s\n", string(body))
	} else {
		fmt.Println("Unsupported content type.")
	}

}
