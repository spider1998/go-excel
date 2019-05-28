package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

//读取表格
func Import(r []byte) (exs [][][]string, sheets []string, lens int, err error) {
	// 打开文件
	xlFile, err := xlsx.OpenBinary(r)
	if err != nil {
		return
	}
	for _, she := range xlFile.Sheets {
		sheets = append(sheets, she.Name)
	}
	results, err := xlFile.ToSlice()
	if err != nil {
		return
	}
	for i, result := range results {
		if len(results[i]) == 0 {
			continue
		}
		exs = append(exs, result[1:])
		lens += len(result[1:])
	}
	return

}

//写入表格
func Export(sheetNames []string, contents [][][]string) (file *xlsx.File, err error) {
	file = xlsx.NewFile()
	for i, content := range contents {
		sheet, err := file.AddSheet(sheetNames[i])
		if err != nil {
			fmt.Printf(err.Error())
		}
		for _, con := range content {
			row := sheet.AddRow()
			row.WriteSlice(&con, len(con))
		}

	}
	return
}
