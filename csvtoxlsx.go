package csvtoxlsx

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"os"

	_ "github.com/xuri/excelize/v2"
)

type CSVToXLSXError struct {
	ErrorCode int
	message   string
}

func (e *CSVToXLSXError) Error() string {
	prefix := fmt.Sprintf("[CSVToXLSX Error: %d]", e.ErrorCode)
	switch e.ErrorCode {
	case -1:
		e.message = "Failed to load CSV file."
	case -2:
		e.message = "Failed to read CSV file."
	default:
		e.message = "Unknown Error"
	}
	return fmt.Sprintf("%s %s", prefix, e.message)
}

func NewCSVToXLSXError(ErrorCode int) *CSVToXLSXError {
	return &CSVToXLSXError{ErrorCode: ErrorCode, message: ""}
}

func LoadCSV(filePath string) ([][]string, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, NewCSVToXLSXError(-1)
	}

	reader := csv.NewReader(bufio.NewReader(file))
	rows, err := reader.ReadAll()
	if err != nil {
		return nil, NewCSVToXLSXError(-2)
	}

	return rows, nil
}

func CreateXLSX(filePath string, sheetName string) {
	// file := excelize.NewFile()
	// index := file.NewSheet(sheetName)
}

func ConvertCSVToXLSX(csvPath string, xlsxPath string) {
	rows, _ := LoadCSV(csvPath)
	for i, row := range rows {
		for k := range row {
			fmt.Printf("%s", rows[i][k])
		}
		fmt.Println("")
	}
}
