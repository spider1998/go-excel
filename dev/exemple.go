package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"strconv"
)

func main() {
	Load()
}

type Person struct {
	Name   string `json:"name"`
	JobNum string `json:"job_num"`
	IDCard string `json:"id_card"`
	Gender int8   `json:"gender"`
	Age    int    `json:"age"`
	Types  int8   `json:"types"`
	Other  string `json:"other"`
}

func Load() (err error) {
	fmt.Println("ready load...")
	var (
		fileName = "./test.xlsx"
		sheet    = "Sheet1"
	)
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	result := xlsx.GetRows(sheet)[1:]
	fmt.Println(len(result))
	for _, list := range result {
		var person Person
		err := SliceToStruct(list, &person)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Println(person)
	}
	return

}

func SliceToStruct(list []string, a interface{}) (err error) {
	v := reflect.ValueOf(a).Elem()
	for i, k := range list {
		sv := v.Field(i).Type().String()
		fmt.Println(sv)
		switch sv {
		case "int":
			intK, _ := strconv.Atoi(k)
			v.Field(i).Set(reflect.ValueOf(intK))
		case "int8":
			intK, _ := strconv.Atoi(k)
			key := int8(intK)
			v.Field(i).Set(reflect.ValueOf(key))
		case "string":
			v.Field(i).Set(reflect.ValueOf(k))
		default:
			panic("invalid type " + v.Field(i).Type().String())
		}
	}
	return
}
