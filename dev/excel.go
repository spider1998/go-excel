package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

var Tag = [27]string{
	"_",
	"A", "B", "C", "D", "E", "F", "G",
	"H", "I", "J", "K", "L", "M", "N",
	"O", "P", "Q", "R", "S", "T",
	"U", "V", "W", "X", "Y", "Z",
}

func main() {
	LoadExcel()
}

func LoadExcel() (err error) {
	fmt.Println("ready load...")
	var (
		fileName   = "./test.xlsx"
		sheet      = "Sheet1"
		startColum = 1
		endColum   = 7
		Insert     = func(list []string) (err error) {
			fmt.Println(list)
			return
		}
	)
	err = StartLoad(fileName, sheet, startColum, endColum, Insert)
	return

}

//读表方法
func StartLoad(fileName, sheet string, startColum, endColum int, Insert func(list []string) (err error)) (err error) {
	var cl string
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	startColum = 1
	startRow := 2
	xlsx.GetRows(sheet)
	endRow := len(xlsx.GetRows(sheet)) - 1
	for i := startRow; i <= endRow; i++ {
		var scales []string
		for j := startColum; j <= endColum; j++ {
			cl = Tag[j] + strconv.Itoa(i)
			cell := xlsx.GetCellValue(sheet, cl)
			scales = append(scales, cell)
		}
		err := Insert(scales)
		if err != nil {
			fmt.Println(err)
		}
	}
	return
}
