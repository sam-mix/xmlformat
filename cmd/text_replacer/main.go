package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	sheetNameRegStrPrefix      = "<Worksheet ss:Name=\""                                               // sheet名称匹配正则表达式字符串前缀
	sheetNameRegStrSuffix      = "\">"                                                                 // sheet名称匹配正则表达式字符串后缀
	sheetNameRegStr            = sheetNameRegStrPrefix + "[^>\x22]+" + sheetNameRegStrSuffix           // sheet名称匹配正则表达式字符串
	stringCellDataRegStrPrefix = "<Data ss:Type=\"String\">"                                           // // 字符串单元格匹配正则表达式字符串前缀
	stringCellDataRegStrSuffix = "</Data>"                                                             // // 字符串单元格匹配正则表达式字符串后缀
	stringCellDataRegStr       = stringCellDataRegStrPrefix + "[^<\x22]+" + stringCellDataRegStrSuffix // 字符串单元格匹配正则表达式字符串
	// oldCfgDataPath             = "./o"                                                                 // 待替换文件所在目录
	// newCfgDataPath             = "./n"                                                                 // 替换完成的文件所在目录
	oldCfgDataPath         = "D:/dev/ws_pub/mol/doc/NewConfig/Config"    // 待替换文件所在目录
	newCfgDataPath         = "D:/dev/ws_pub/mol/doc/NewConfig/NewConfig" // 替换完成的文件所在目录
	textReplaceRulesFile   = "text_replace_rules.xlsx"                   // 文本替换规则配置文件
	sheetText              = "文本替换"                                      // 文本替换规则配置文件中文本替换sheet名
	sheetCell              = "单元替换"                                      // 文本替换规则配置文件中单元替换sheet名
	sheetSheetNameContains = "sheet名称部分替换"                               // 文本替换规则配置文件中sheet名称部分替换sheet名
	sheetSheetName         = "sheet名称完整替换"                               // 文本替换规则配置文件中sheet名称完整替换sheet名
	cellSplitSep           = ">"
	nameSolitSep           = "\""
	isPart                 = 0
	isAll                  = 1
	cfgDataFilesuffix      = "_data.xml"
)

// 规则 rule
type r struct {
	n string              // 替换成的内容
	l int                 // 在替换规则文件中的行号
	w map[string]struct{} // 配置文件白名单
	b map[string]struct{} // 配置文件黑名单
}

// 错误规则 error rule
type er struct {
	o string // 待替换的内容
	l int    // 在替换规则文件中第一次出现的行号
}

var (
	rulesSheetNameList = [4]string{sheetText, sheetCell, sheetSheetNameContains, sheetSheetName} // 替换规则配置文件sheet名列表
	rules              = make(map[string]map[string]r)                                           // map[替换规则文件中的sheet名称]map[待替换内容]替换规则
	wg                 sync.WaitGroup                                                            // 同步机制 等待组
	allXmlPath         = make([]string, 0)
	cellReg            = regexp.MustCompile(stringCellDataRegStr)
	sheetNameReg       = regexp.MustCompile(sheetNameRegStr)
)

func main() {

	s := time.Now()
	// 获取替换规则
	getReplaceRules()
	// 替换全部
	replaceAll()

	e := time.Since(s)
	fmt.Printf("替换完成!共耗时 %f 秒 \n", e.Seconds())
}

func replaceAll() {
	// 获取所有xml路径
	getAllXmlFilePath()
	wg.Add(len(allXmlPath))
	for _, filePath := range allXmlPath {
		go replace(filePath, &wg)
	}
	wg.Wait()
}

// 替换
func replace(filePath string, waitGroup *sync.WaitGroup) {
	if len(rules) > 0 {
		content := readFile(filePath)
		_, fileName := filepath.Split(filePath)
		shortFileName := strings.ReplaceAll(fileName, cfgDataFilesuffix, "")
		content, countName := replaceSheetName(shortFileName, content)
		content, countText := replaceText(shortFileName, content)
		if (countName + countText) > 0 {
			ioutil.WriteFile(strings.ReplaceAll(filePath, oldCfgDataPath, newCfgDataPath), []byte(content), 0644)
			fmt.Printf("文件: %50s sheet name 替换 %6d 次，cell text 替换 %6d 次\n", fileName, countName, countText)

		}
	}
	defer wg.Done()
}

// 替换cell文本
func replaceText(shortFileName, content string) (string, int) {
	return replaceCommon(content, sheetText, sheetCell, sheetNameRegStrPrefix, sheetNameRegStrSuffix, shortFileName, cellSplitSep, cellReg)
}

// 通用替换
func replaceCommon(content, sheetText, sheetCell, sheetNameRegStrPrefix, sheetNameRegStrSuffix, shortFileName, splitSep string, regexp *regexp.Regexp) (string, int) {
	rulesText, okT := rules[sheetText]
	rulesCell, okC := rules[sheetCell]
	if !okC && !okT {
		return content, 0
	}
	filterRulesText := filterRules(okT, rulesText, shortFileName)
	filterRulesCell := filterRules(okC, rulesCell, shortFileName)
	if (len(filterRulesCell) + len(filterRulesText)) < 0 {
		return content, 0
	}
	ress := regexp.FindAllString(content, -1)
	count := 0
	content, count = textReplace(filterRulesText, sheetNameRegStrPrefix, sheetNameRegStrSuffix, content, splitSep, ress, count, isPart)
	return textReplace(filterRulesCell, sheetNameRegStrPrefix, sheetNameRegStrSuffix, content, splitSep, ress, count, isAll)
}

// 过滤条件
func filterRules(okC bool, rules map[string]r, shortFileName string) map[string]r {
	filterRules := make(map[string]r)
	if okC {
		for o, r := range rules {
			_, okw := r.w[shortFileName]
			_, okb := r.b[shortFileName]
			if (len(r.w) == 0 || okw) && !okb {
				filterRules[o] = r
			}
		}
	}
	return filterRules
}

// 文本替换
func textReplace(rules map[string]r, regStrPrefix, regStrSuffix, content, splitSep string, ress []string, count int, optype int) (string, int) {
	if len(rules) > 0 {
		for _, res := range ress {
			os := strings.Split(res, splitSep)[1]
			resTrim := strings.TrimSpace(os)
			for o, r := range rules {
				if (optype == isAll && resTrim == o) || (optype == isPart && strings.Contains(resTrim, o)) {
					nr := strings.ReplaceAll(resTrim, o, r.n)
					anr := regStrPrefix + nr + regStrSuffix
					content = strings.ReplaceAll(content, res, anr)
					count += 1
				}
			}

		}
	}
	return content, count
}

// 替换sheet name
func replaceSheetName(shortFileName, content string) (string, int) {
	return replaceCommon(content, sheetSheetNameContains, sheetSheetName, sheetNameRegStrPrefix, sheetNameRegStrSuffix, shortFileName, nameSolitSep, sheetNameReg)
}

// 读取文件
func readFile(filePath string) string {
	bytes, err := ioutil.ReadFile(filePath)
	if err != nil {
		panic(err)
	}
	return string(bytes)
}

// 获取所有xml路径
func getAllXmlFilePath() {
	allXmlPath = listDir(oldCfgDataPath, cfgDataFilesuffix)
}

//获取指定目录下的所有文件，不进入下一级目录搜索，可以匹配后缀过滤。
func listDir(dirPth string, suffix string) []string {
	files := make([]string, 0)

	dir, err := ioutil.ReadDir(dirPth)
	if err != nil {
		panic(err)
	}

	PthSep := string(os.PathSeparator)
	suffix = strings.ToUpper(suffix) //忽略后缀匹配的大小写

	for _, fi := range dir {
		if fi.IsDir() { // 忽略目录
			continue
		}
		if strings.HasSuffix(strings.ToUpper(fi.Name()), suffix) { //匹配文件
			files = append(files, dirPth+PthSep+fi.Name())
		}
	}

	return files
}

// 获取替换规则
func getReplaceRules() {
	f, err := excelize.OpenFile(textReplaceRulesFile)
	if err != nil {
		panic(err)
	}
	errMap := make(map[string]map[int]er) // map[替换规则文件中的sheet名称]map[重复的待替换内容所在位置]er
	for _, sheet := range rulesSheetNameList {
		m := make(map[string]r)
		em := make(map[int]er)
		rows, err := f.GetRows(sheet)
		if err != nil {
			panic(err)
		}
		for i, row := range rows {
			o, n, ws, bs := "", "", "", ""
			rowLen := len(row)
			if rowLen > 0 {
				o = strings.TrimSpace(row[0])
			}
			if o == "" {
				continue
			}
			if fr, ok := m[o]; ok {
				em[i+1] = er{
					o: o,
					l: fr.l,
				}
				continue
			}

			if rowLen > 1 {
				n = strings.TrimSpace(row[1])
			}
			if rowLen > 2 {
				ws = strings.TrimSpace(row[2])
			}
			if rowLen > 3 {
				bs = strings.TrimSpace(row[3])
			}
			wss := strings.Split(ws, ",")
			bss := strings.Split(bs, ",")
			if wss[0] == "" {
				wss = wss[1:]
			}
			if bss[0] == "" {
				bss = bss[1:]
			}

			m[o] = r{
				n: n,
				w: l2m(wss),
				b: l2m(bss),
				l: i + 1,
			}
		}
		if len(m) > 0 {
			rules[sheet] = m
		}
		if len(em) > 0 {
			errMap[sheet] = em
		}
	}
	if len(errMap) > 0 {
		fmt.Printf("替换规则配置文件(%s)存在警告 (最早定义的将会生效，后定义的将不会生效):\n", textReplaceRulesFile)
		for sheet, em := range errMap {
			fmt.Printf("工作区(sheet): %s :\n", sheet)
			for line, e := range em {
				fmt.Printf("行: %d, 定义的待替换内容: %s, 已经在%d行定义\n", line, e.o, e.l)
			}
		}
	}

}

func l2m(l []string) map[string]struct{} {
	m := make(map[string]struct{})
	for _, k := range l {
		m[k] = struct{}{}
	}
	return m
}
