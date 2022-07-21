package csvtoxlsx

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func LoadCSV(csvPath string) ([][]string, error) {
	var csvData [][]string

	csvFile, err := os.Open(csvPath)
	if err != nil {
		return [][]string{}, err
	}
	defer func() {
		if err = csvFile.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	csvReader := csv.NewReader(csvFile)
	csvData, err = csvReader.ReadAll()
	if err != nil {
		return [][]string{}, err
	}

	return csvData, err
}

func ConvertCSVToXLSX(csvPath string, xlsxPath string, sheetName string) error {
	xlsxPathAbs, _ := filepath.Abs(xlsxPath)
	_, xlsxFileName := filepath.Split(xlsxPathAbs)

	var csvData [][]string
	var err error

	csvData, err = LoadCSV(csvPath)
	if err != nil {
		return err
	}

	xlsxFile := excelize.NewFile()
	xlsxFile.NewSheet(sheetName)
	defer func() {
		if err = xlsxFile.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	var lastCellNum string
	if len(csvData) == 0 || len(csvData[0]) == 0 {
		lastCellNum = "$A$1"
	} else {
		lastCellNum, _ = excelize.CoordinatesToCellName(len(csvData[0]), len(csvData), true)
	}

	for i, row := range csvData {
		for k, col := range row {
			cellNum, _ := excelize.CoordinatesToCellName(k+1, i+1)
			err := xlsxFile.SetCellValue(sheetName, cellNum, col)
			if err != nil {
				return err
			}
		}
	}

	xlsxFileNameWithoutExt := xlsxFileName[:strings.LastIndex(xlsxFileName, filepath.Ext(xlsxFileName))]
	xlsxFile.SetDefinedName(&excelize.DefinedName{
		Name:     fmt.Sprintf("%s %s", time.Now().Format("20060102"), xlsxFileNameWithoutExt),
		RefersTo: fmt.Sprintf("%s!$A$1:%s", sheetName, lastCellNum),
		Scope:    sheetName,
	})

	err = xlsxFile.SaveAs(xlsxPathAbs, excelize.Options{
		RawCellValue: true,
	})
	if err != nil {
		return err
	}

	return nil
}
