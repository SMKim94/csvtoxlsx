package csvtoxlsx

import (
	"fmt"
	"log"
	"testing"

	"github.com/xuri/excelize/v2"
)

var csvPath string = "./0.csv"

func TestLoadCSV(t *testing.T) {
	filePath := csvPath
	_, err := LoadCSV(filePath)
	if err != nil {
		fmt.Println(err.Error())
		t.Fatalf(err.Error())
	}

	// for i, row := range rows {
	// 	for k := range row {
	// 		fmt.Printf("%s", rows[i][k])
	// 	}
	// 	fmt.Println("")
	// }
}

func TestConvertCSVToXLSX(t *testing.T) {
	err := ConvertCSVToXLSX(csvPath, "./", "myExcel", "test")
	if err != nil {
		t.Fatalf(err.Error())
	}
}

func TestSaveXLSX(t *testing.T) {
	file := excelize.NewFile()
	sheetIdx := file.NewSheet("Sheet1")
	log.Printf("sheetIdx: %d\n", sheetIdx)
	defer file.Close()

	for i := 0; i < 100; i++ {
		for k := 0; k < 50; k++ {
			cellNum, err := excelize.CoordinatesToCellName(k+1, i+1)
			if err != nil {
				t.Fatalf(err.Error())
			}

			err = file.SetCellValue("Sheet1", cellNum, fmt.Sprintf("%s:%d", cellNum, i*k))
			if err != nil {
				t.Fatalf(err.Error())
			}
		}
	}

	err := file.SetDefinedName(&excelize.DefinedName{
		Name:     "My Excel File test",
		RefersTo: "Sheet1!$A$1:$A$2",
		Comment:  "This is comment.",
		Scope:    "Sheet1",
	})
	if err != nil {
		t.Fatalf(err.Error())
	}

	err = file.SaveAs("Test.xlsx")
	if err != nil {
		t.Fatalf(err.Error())
	}

	err = file.Close()
	if err != nil {
		t.Fatalf(err.Error())
	}
}
