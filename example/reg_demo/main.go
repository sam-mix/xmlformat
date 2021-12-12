package main

import (
	"fmt"
	"io/ioutil"
	"regexp"
)

const (
	sheetNameRegStr      = "<Worksheet ss:Name=\"[^>\x22]+\">"         // sheet名称匹配正则表达式字符串
	stringCellDataRegStr = "<Data ss:Type=\"String\">[^<\x22]+</Data>" // 字符串单元格匹配正则表达式字符串
	oldCfgDataPath       = "D:/dev/ws_pub/mol/doc/NewConfig/Config/item_data.xml"
	new
)

func main() {
	reg := regexp.MustCompile(stringCellDataRegStr)

	bytes, err := ioutil.ReadFile("")
	if err != nil {
		panic(err)
	}
	res := reg.FindAllString(string(bytes), -1)
	for _, r := range res {
		fmt.Println(r)
	}
}
