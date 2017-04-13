// Copyright 2016 zxfonline@sina.com. All rights reserved.
// Use of this source code is governed by a BSD-style
// license that can be found in the LICENSE file.

package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"path"
	"strings"

	"github.com/zxfonline/fileutil"

	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
var delimiter = flag.String("d", ",", "Delimiter to use between fields")
var csvPath = flag.String("o", "", "Path to the CSV output file")

type outputer func(s string)

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int, outputf outputer) error {
	excelFileName = strings.Replace(excelFileName, "\\", "/", -1)
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		return error
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets.")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d\n", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]
	for _, row := range sheet.Rows {
		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
				str, err := cell.String()
				if err != nil {
					//					vals = append(vals, err.Error())
					return err
				}
				if strings.ContainsAny(str, `"`) {
					str = strings.Replace(str, `"`, `""`, -1)
					//					vals = append(vals, fmt.Sprintf("\"%s\"", str))
					//				} else if strings.ContainsAny(str, "\n") {
					//					vals = append(vals, fmt.Sprintf("\"%s\"", str))
					//				} else {
					//					vals = append(vals, str)
					//					vals = append(vals, fmt.Sprintf("\"%s\"", str))
				}
				vals = append(vals, fmt.Sprintf("\"%s\"", str))
			}
			outputf(strings.Join(vals, *delimiter) + "\n")
		}
	}
	return nil
}

//构建一个每日写日志文件的写入器
func openFile(pathfile string) (wc *os.File, err error) {
	dir, _ := path.Split(pathfile)
	if _, err = os.Stat(dir); err != nil && !os.IsExist(err) {
		if !os.IsNotExist(err) {
			return nil, err
		}
		if err = os.MkdirAll(dir, os.ModePerm); err != nil {
			return nil, err
		}
		if _, err = os.Stat(dir); err != nil {
			return nil, err
		}
	}
	return os.OpenFile(pathfile, os.O_TRUNC|os.O_CREATE|os.O_WRONLY, os.ModePerm)
}

func main() {
	flag.Parse()
	if *xlsxPath == "" {
		if len(os.Args) < 4 {
			flag.PrintDefaults()
			return
		}
		flag.Parse()
	}
	if *csvPath == "" {
		*csvPath = fileutil.ChangeFileExt(*xlsxPath, ".csv")
	}
	wc, err := openFile(*csvPath)
	if err != nil {
		panic(err)
	}
	defer wc.Close()
	printer := func(s string) {
		//		fmt.Printf("%s", s)
		if _, err := wc.WriteString(s); err != nil {
			panic(err)
		}

	}
	if err := generateCSVFromXLSXFile(*xlsxPath, *sheetIndex, printer); err != nil {
		panic(err)
	}
}
