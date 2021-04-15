package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/beevik/etree"
)

func main() {
	input := "Data/battle/Battle.xlsx"
	f, err := excelize.OpenFile(input)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get value from cell by given worksheet name and axis.
	// cell, err := f.GetCellValue("Sheet1", "B2")
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }
	// fmt.Println(cell)
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Battle")
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
	// 直接生成lua
	// 直接生成csv
	// 直接生成xml描述表
	writeCSV(f)
	writeLua(f)
	writeXML(f)
}

func writeLua(f *excelize.File) error {
	rows, err := f.GetRows("Battle")
	if err != nil {
		return err
	}

	var file *os.File
	var err1 error
	filename := "Battle.lua"
	if checkFileIsExist(filename) { //如果文件存在
		// file, err1 = os.OpenFile(filename, os.O_APPEND, 0666) //打开文件
		fmt.Println("文件存在")
		os.Remove(filename)
	} else {
		// file, err1 = os.Create(filename) //创建文件
		fmt.Println("文件不存在")
	}

	file, err1 = os.Create(filename) //创建文件
	if err1 != nil {
		return err1
	}

	col_name := make([]string, 0)
	col_type := make([]string, 0)
	writer := bufio.NewWriter(file)
	writer.WriteString("local battle = {\n")
	for i, row := range rows {
		if i <= 2 {
			if i == 1 {
				col_name = append(col_name, row...)
				continue
			}

			if i == 2 {
				col_type = append(col_type, row...)
				continue
			}
			continue
		}

		val := ""
		key := ""
		for k, colCell := range row {
			if k != 0 {
				val += ","
			} else {
				key = colCell
			}

			if col_type[k] == "STRING" {
				if strings.Contains(colCell, "'") {
					val += fmt.Sprintf("%s = \"%s\"", col_name[k], colCell)
				} else {
					val += fmt.Sprintf("%s = '%s'", col_name[k], colCell)
				}

			} else {
				val += fmt.Sprintf("%s = %s", col_name[k], colCell)
			}
		}

		if key == "" || val == "" {
			continue
		}

		val = fmt.Sprintf("\t[%s] = {%s},", key, val)
		if i != 3 {
			writer.WriteString("\n")
		}
		writer.WriteString(val)
	}

	writer.WriteString("\n}")
	writer.WriteString("\nreturn battle")

	writer.WriteString("\n\n---@class TableBattle")
	for i := 0; i < len(col_name); i++ {
		typename := "number"
		if col_type[i] == "STRING" {
			typename = "string"
		}
		val := fmt.Sprintf("\n---@field %s %s", col_name[i], typename)
		writer.WriteString(val)
	}

	writer.Flush()
	file.Close()
	return nil
}

func writeCSV(f *excelize.File) error {
	rows, err := f.GetRows("Battle")
	if err != nil {
		return err
	}

	var file *os.File
	var err1 error
	filename := "Battle.csv"
	if checkFileIsExist(filename) { //如果文件存在
		// file, err1 = os.OpenFile(filename, os.O_APPEND, 0666) //打开文件
		fmt.Println("文件存在")
		os.Remove(filename)
	} else {
		// file, err1 = os.Create(filename) //创建文件
		fmt.Println("文件不存在")
	}

	file, err1 = os.Create(filename) //创建文件
	if err1 != nil {
		return err1
	}

	writer := bufio.NewWriter(file)
	writer.WriteString("\xEF\xBB\xBF")

	for i, row := range rows {
		if i != 0 {
			writer.WriteString("\n")
		}

		for k, colCell := range row {
			// fmt.Print(colCell, "\t")
			if k != 0 {
				writer.WriteString(",")
			}
			if strings.Contains(colCell, ",") {
				writer.WriteString(fmt.Sprintf("\"%s\"", colCell))
			} else {
				writer.WriteString(colCell)
			}

		}
		// fmt.Println()

	}
	writer.Flush()
	file.Close()
	return nil
}

func writeXML(f *excelize.File) error {
	rows, err := f.GetRows("Battle")
	if err != nil {
		return err
	}

	col_name := make([]string, 0)
	col_type := make([]string, 0)
	for i, row := range rows {
		if i <= 2 {
			if i == 1 {
				col_name = append(col_name, row...)
				continue
			}

			if i == 2 {
				col_type = append(col_type, row...)
				continue
			}
			continue
		}
	}

	filename := "Description.xml"
	doc := etree.NewDocument()
	doc.CreateProcInst("xml", `version="1.0" encoding="UTF-8"`)
	application := doc.CreateElement("DataPool")
	application.CreateAttr("project", "game")
	application.CreateAttr("version", "1.0.0")

	tableNode := application.CreateElement("DataElement")
	tableNode.CreateAttr("Name", "Battle")
	for i := 0; i < len(col_name); i++ {
		fieldNode := tableNode.CreateElement("Field")
		fieldNode.CreateAttr("Name", col_name[i])
		typename := "string"
		if col_type[i] == "INT" {
			typename = "int"
		}
		fieldNode.CreateAttr("Type", typename)
	}

	doc.Indent(2)
	doc.WriteToFile(filename)
	return nil
}

/**
 * 判断文件是否存在  存在返回 true 不存在返回false
 */
func checkFileIsExist(filename string) bool {
	var exist = true
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		exist = false
	}
	return exist
}
