package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"strconv"

	"flag"

	"path/filepath"

	"github.com/tealeg/xlsx"
)

type XlBook struct {
	Sheets []XlSheet
}
type XlSheet struct {
	Name string
	Data []XlRow
}
type XlRow struct {
	RowNumber int
	Cells     []XlCell
}
type XlCell struct {
	Type    xlsx.CellType
	Formula string
	Value   string
}

func xlsx2json(input string, output string) {
	excelFileName := input
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
	}
	var xlBook XlBook
	for _, sheet := range xlFile.Sheets {
		var xlSheet XlSheet
		xlSheet.Name = sheet.Name
		for i, row := range sheet.Rows {
			var xlRow XlRow
			xlRow.RowNumber = i
			for _, cell := range row.Cells {
				xlRow.Cells = append(xlRow.Cells, XlCell{
					Type:    cell.Type(),
					Formula: cell.Formula(),
					Value:   cell.Value,
				})
			}
			xlSheet.Data = append(xlSheet.Data, xlRow)
		}
		xlBook.Sheets = append(xlBook.Sheets, xlSheet)
	}
	jsonData, _ := json.MarshalIndent(xlBook, "", "\t")
	f, err := os.Create(output)
	if err != nil {
		log.Fatal(err)
	}
	if _, err := f.Write(jsonData); err != nil {
		log.Fatal(err)
	}
	defer f.Close()
}

func json2xlsx(input string, output string) {
	var err error
	jsonFileName := input
	jsonFile, err := ioutil.ReadFile(jsonFileName)
	if err != nil {
		log.Fatal(err)
	}
	var xlBook XlBook
	err = json.Unmarshal(jsonFile, &xlBook)
	if err != nil {
		log.Fatal(err)
	}
	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	file = xlsx.NewFile()
	for _, xlSheet := range xlBook.Sheets {
		sheet, err = file.AddSheet(xlSheet.Name)
		if err != nil {
			log.Fatal(err)
		}
		for _, xlRow := range xlSheet.Data {
			row = sheet.AddRow()
			for _, xlCell := range xlRow.Cells {
				cell = row.AddCell()
				switch xlCell.Type {
				case 0:
					cell.SetString(xlCell.Value)
				case 1:
					cell.SetFormula(xlCell.Formula)
				case 2:
					v, err := strconv.ParseFloat(xlCell.Value, 64)
					if err != nil {
						log.Fatal(err)
					}
					cell.SetFloat(v)
				case 3:
					v, err := strconv.ParseBool(xlCell.Value)
					if err != nil {
						log.Fatal(err)
					}
					cell.SetBool(v)
				case 6:
					v, err := strconv.ParseFloat(xlCell.Value, 64)
					if err != nil {
						log.Fatal(err)
					}
					cell.SetDate(xlsx.TimeFromExcelTime(v, false))
				default:
					if xlCell.Formula == "" {
						cell.SetString(xlCell.Value)
					} else {
						cell.SetFormula(xlCell.Formula)
					}
				}
			}
		}
	}
	err = file.Save(output)
	if err != nil {
		log.Fatal(err)
	}
}

func main() {
	var input, output string
	flag.StringVar(&input, "input", "", "input file name (*.xlsx or *.json)")
	flag.StringVar(&output, "output", "", "output file name")
	flag.Parse()
	if output == "" {
		if filepath.Ext(input) == ".xlsx" {
			output = filepath.Base(input) + ".json"
		} else {
			output = filepath.Base(input) + ".xlsx"
		}
	}
	if filepath.Ext(input) == ".xlsx" && filepath.Ext(output) == ".json" {
		xlsx2json(input, output)
	} else if filepath.Ext(input) == ".json" && filepath.Ext(output) == ".xlsx" {
		json2xlsx(input, output)
	} else {
		fmt.Println("Not *.xlsx <=> *.json, bye!")
		os.Exit(1)
	}
}
