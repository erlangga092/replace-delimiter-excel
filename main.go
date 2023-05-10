package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/alecthomas/kingpin"
	"github.com/tealeg/xlsx"
)

var (
	app           = kingpin.New("App", "Remove dot from excel file")
	argFile       = app.Flag("file", "Excel file").Short('f').Required().String()
	argWorksheet  = app.Flag("worksheet", "Worksheet").Short('w').Required().String()
	argStartRow   = app.Flag("start-row", "Start row index").Short('s').Required().Int()
	argEndRow     = app.Flag("end-row", "Last row index").Short('e').Required().Int()
	argFileOutput = app.Flag("file output", "File output").Short('o').Required().String()
	argDelimiter  = app.Flag("delimiter", "Delimiter").Short('d').Required().String()
)

func main() {
	kingpin.MustParse(app.Parse(os.Args[1:]))

	xlFile, err := xlsx.OpenFile(*argFile)
	if err != nil {
		panic(err)
	}

	sheet := xlFile.Sheet[*argWorksheet]

	startRow := strconv.Itoa(*argStartRow)
	startRowInt, _ := strconv.Atoi(startRow)

	endRow := strconv.Itoa(*argEndRow)
	endRowInt, _ := strconv.Atoi(endRow)

	for rowIndex := startRowInt; rowIndex < endRowInt; rowIndex++ { // Baris 2 hingga 141
		cell := sheet.Cell(rowIndex, 2)
		// if cell.Type() == xlsx.CellTypeNumeric { // Pastikan tipe sel adalah angka
		if cell.Type() == 0 {
			value := cell.String()
			newStr := strings.Replace(value, *argDelimiter, "", -1)
			fmt.Println(newStr)
			cell.SetValue(newStr)
		} else {
			fmt.Println("bukan angka")
		}
	}

	err = xlFile.Save(*argFileOutput)
	if err != nil {
		panic(err)
	}

	fmt.Println("File berhasil disimpan.")
}
