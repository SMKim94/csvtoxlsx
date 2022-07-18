package csvtoxlsx

import (
	"testing"
	"fmt"
)

func TestLoadCSV(t *testing.T) {
	filePath := "./0.csv"
	rows, err := LoadCSV(filePath)
	if err != nil {
		fmt.Println(err.Error())
		t.Fatalf(err.Error())
	}

	for i, row := range rows {
		for k := range row {
			fmt.Printf("%s", rows[i][k])
		}
		fmt.Println("")
	}
}