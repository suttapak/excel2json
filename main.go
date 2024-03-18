package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	start := time.Now()
	defer func() {
		fmt.Println("time: ", time.Since(start))
	}()
	filePaht := flag.String("f", "", "file path")
	outDest := flag.String("o", "", "output destination")
	sheetName := flag.String("s", "", "sheet name")

	flag.Parse()

	if *filePaht == "" {
		flag.PrintDefaults()
		return
	}
	if *outDest == "" {
		flag.PrintDefaults()
		return
	}
	if *sheetName == "" {
		flag.PrintDefaults()
		return
	}

	filePathStr, err := filepath.Abs(*filePaht)
	if err != nil {
		panic(err)

	}

	f, err := excelize.OpenFile(filePathStr)
	if err != nil {
		panic(err)

	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			panic(err)

		}
	}()

	rows, err := f.GetRows(*sheetName)
	if err != nil {
		panic(err)
	}
	var key map[int]string
	var value []map[string]interface{}
	for i, row := range rows {
		if i == 0 {
			key = make(map[int]string)
			for j, colCell := range row {
				key[j] = colCell
			}
			continue
		}

		if len(key) != len(row) {
			for i := len(row); i < len(key); i++ {
				row = append(row, "")
			}

		}
		// map value key to json
		r := make(map[string]interface{})
		for index, colCell := range row {

			if key[index] == ("เงื่อนไขการรับประกัน") {
				if colCell == "" {
					colCell = "1"
				}
				if len(colCell) > 0 {
					colCell = strings.Split(colCell, " ")[0]
				}
				colCell, err := strconv.Atoi(colCell)
				if err != nil {
					fmt.Println(err)
					return
				}
				r[key[index]] = colCell
				continue
			}
			r[key[index]] = strings.Trim(colCell, " ")

		}
		value = append(value, r)
	}
	// value to json file
	j, _ := json.Marshal(value)
	outPaht, err := filepath.Abs("dist/" + *outDest)
	if err != nil {
		panic(err)
	}

	err = os.WriteFile(outPaht, j, 0644)
	if err != nil {
		panic(err)
	}
}
