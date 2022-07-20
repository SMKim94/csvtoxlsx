package csvtoxlsx

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"

	"github.com/xuri/excelize/v2"
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

func ConvertCSVToXLSX(csvPath string, xlsxPath string, sheetName string) error {
	csvPathAbs, _ := filepath.Abs(csvPath)
	csvDir, csvFileName := filepath.Split(csvPathAbs)
	xlsxPathAbs, _ := filepath.Abs(xlsxPath)
	xlsxDir, xlsxFileName := filepath.Split(xlsxPathAbs)

	fmt.Printf("CSV Dir: %s\n", csvDir)
	fmt.Printf("CSV FileName: %s\n", csvFileName)
	fmt.Printf("XLSX Dir: %s\n", xlsxDir)
	fmt.Printf("XLSX FileName: %s\n", xlsxFileName)

	// ----------------------------------------------
	var csvData [][]string

	csvFile, err := os.Open(csvPath)
	if err != nil {
		return err
	}
	defer func() {
		if err = csvFile.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	csvReader := csv.NewReader(csvFile)
	csvData, err = csvReader.ReadAll()
	if err != nil {
		return err
	}
	// for {
	// 	rec, err := csvReader.Read()
	// 	if err == io.EOF {
	// 		break
	// 	}
	// 	if err != nil {
	// 		return [][]string{}, err
	// 	}

	// 	fmt.Printf("%+v\n", rec)
	// 	csvData = append(csvData, rec)
	// }
	// fmt.Printf("%+v\n", csvData)
	// ----------------------------------------------

	// ----------------------------------------------
	xlsxFile := excelize.NewFile()
	xlsxFile.NewSheet(sheetName)
	defer func() {
		if err = xlsxFile.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// var lastCellNum string
	// if len(csvData) == 0 || len(csvData[0]) == 0 {
	// 	lastCellNum = "A1"
	// } else {
	// 	lastCellNum, _ = excelize.CoordinatesToCellName(len(csvData[0]), len(csvData))
	// }

	for i, row := range csvData {
		for k, col := range row {
			cellNum, _ := excelize.CoordinatesToCellName(k+1, i+1)
			fmt.Printf("{%s: %v}\t", cellNum, col)

			err := xlsxFile.SetCellValue(sheetName, cellNum, col)
			if err != nil {
				return err
			}
		}
		fmt.Println("")
	}

	err = xlsxFile.SaveAs(xlsxPathAbs, excelize.Options{
		RawCellValue: true,
	})
	if err != nil {
		return err
	}

	// ----------------------------------------------

	return nil
}

func ConvertCSVToXLSX_0(csvPath string, xlsxPath string, xlsxInfo ...string) error {
	regPath := regexp.MustCompile(`.{1,}\/`)
	regExt := regexp.MustCompile(`\.[c,C][s,S][v,V]`)
	csvFileName := regExt.ReplaceAllString(regPath.ReplaceAllString(csvPath, ""), "")
	xlsxFileName := csvFileName
	xlsxSheetName := csvFileName

	switch len(xlsxInfo) {
	case 1:
		xlsxFileName = xlsxInfo[0]
		fallthrough
	case 2:
		xlsxSheetName = xlsxInfo[1]
	}

	fmt.Printf("File: %s, Sheet: %s\n", xlsxFileName, xlsxSheetName)
	xlsx := excelize.NewFile()
	xlsx.NewSheet(xlsxSheetName)
	defer xlsx.Close()

	// CSV 파일 데이터 불러오기
	rows, err := LoadCSV(csvPath)
	if err != nil {
		return err
	}

	var lastCellNum string
	// XLSX 파일에 CSV 파일 데이터 넣기
	for i, row := range rows {
		for k := range row {
			cellNum, cellErr := excelize.CoordinatesToCellName(k+1, i+1)
			if cellErr != nil {
				return cellErr
			}
			log.Printf("%s, %s, %s", xlsxSheetName, cellNum, rows[i][k])
			if err := xlsx.SetCellValue(xlsxSheetName, cellNum, rows[i][k]); err != nil {
				return err
			}

			if lastRow, lastCol := len(rows), len(rows[i]); i == lastRow && k == lastCol {
				lastCellNum, cellErr = excelize.CoordinatesToCellName(lastCol, lastRow, true)
				log.Printf("lastCellNum: %s", lastCellNum)
				if cellErr != nil {
					return cellErr
				}
			}
		}
	}

	xlsxErr := xlsx.SetDefinedName(&excelize.DefinedName{
		Name:     xlsxFileName,
		RefersTo: fmt.Sprintf("%s!$A$1:%s", xlsxSheetName, lastCellNum),
		Comment:  "XLSX file created by csvtoxlsx.",
		Scope:    xlsxSheetName,
	})
	if xlsxErr != nil {
		return xlsxErr
	}

	xlsxErr = xlsx.SaveAs(fmt.Sprintf("%s.xlsx", xlsxFileName))
	if xlsxErr != nil {
		return xlsxErr
	}

	xlsxErr = xlsx.Close()
	if xlsxErr != nil {
		return xlsxErr
	}

	return xlsxErr
}
