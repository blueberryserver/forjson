package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"strconv"

	"github.com/tealeg/xlsx"
)

func main() {
	fmt.Println("for json")

	fileType := flag.String("type", "json", "TYPE")
	split := flag.String("split", "no", "SPLIT")

	flag.Parse()
	//*
	//xlsx 파일 오픈
	file, err := xlsx.OpenFile("SimTable.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	if *fileType == "json" {
		if *split == "single" {
			createJSONSingle("SimTable", file)
		} else {
			for _, sheet := range file.Sheets {
				tableName, jsonStr := createJSON(sheet)
				file, err := os.Create(fmt.Sprintf("%s.json", tableName))
				if err != nil {
					fmt.Println(err)
					return
				}
				file.WriteString(jsonStr)
				file.Close()
			}
		}
	} else if *fileType == "csv" {
		for _, sheet := range file.Sheets {
			createCSV(sheet)
		}
	} else if *fileType == "proto" {
		createProto("SimTable", "MSG", file)
	}

	//파일 트레이싱
	// for s, sheet := range file.Sheets {
	// 	for r, row := range sheet.Rows {
	// 		for c, cell := range row.Cells {
	// 			text := cell.String()
	// 			fmt.Printf("%d %d %d  %s\n", s, r, c, text)
	// 		}
	// 	}
	// }
	/**/
	// 엑셀 파싱
	// json 파일 생성
}

func createJSON(sheet *xlsx.Sheet) (string, string) {
	// 0, 0 테이블명
	tableName := sheet.Rows[1].Cells[0].String()
	//fmt.Println(tableName)
	js := "{\"TableName\":\"" + tableName + "\",\r\n"

	cellcount := len(sheet.Rows[6].Cells)
	_ = len(sheet.Rows) - 7
	//fmt.Println(cellcount, rowcount)

	types := make([]string, cellcount)
	for i, cell := range sheet.Rows[4].Cells {
		types[i] = cell.String()
	}
	//fmt.Println(types)

	columns := make([]string, cellcount)
	for i, cell := range sheet.Rows[6].Cells {
		columns[i] = cell.String()
	}

	data := "\"Data\":["
	for r, row := range sheet.Rows {
		if r < 7 {
			continue
		}

		jsonStr := "{"
		for i, cell := range row.Cells {
			jsonStr = jsonStr + "\"" + columns[i] + "\":"
			if types[i] == "String" {
				jsonStr = jsonStr + "\"" + cell.String() + "\""
			} else {
				if len(cell.String()) == 0 {
					jsonStr = jsonStr + "0"
				} else {
					jsonStr = jsonStr + cell.String()
				}
			}

			if i < cellcount-1 {
				jsonStr = jsonStr + ","
			}
		}
		jsonStr = jsonStr + "}\r\n"
		//fmt.Println(jsonStr)
		data = data + jsonStr

		if r < len(sheet.Rows)-1 {
			data = data + ","
		}
	}
	data = data + "]"

	js = js + data + "}"

	return tableName, js
	//fmt.Println(js)
}

//
func createJSONSingle(fileName string, file *xlsx.File) {
	js := "{\r\n"
	for i, sheet := range file.Sheets {
		js = js + "\"" + sheet.Name + "\":"
		_, jsonStr := createJSON(sheet)
		js = js + jsonStr

		if i < len(file.Sheets)-1 {
			js = js + ",\r\n"
		}
	}
	js = js + "}"

	outFile, err := os.Create(fmt.Sprintf("%s.json", fileName))
	if err != nil {
		fmt.Println(err)
		return
	}
	outFile.WriteString(js)
	outFile.Close()
}

func createCSV(sheet *xlsx.Sheet) {
	tableName := sheet.Rows[1].Cells[0].String()

	csvfile, err := os.Create(fmt.Sprintf("%s.csv", tableName))
	defer csvfile.Close()
	if err != nil {
		return
	}

	cellcount := len(sheet.Rows[6].Cells)
	_ = len(sheet.Rows) - 7

	types := make([]string, cellcount)
	for i, cell := range sheet.Rows[4].Cells {
		types[i] = cell.String()
	}

	columns := make([]string, cellcount)
	for i, cell := range sheet.Rows[6].Cells {
		columns[i] = cell.String()
	}

	csvwriter := csv.NewWriter(csvfile)
	csvwriter.Write(columns)

	for r, row := range sheet.Rows {
		if r < 7 {
			continue
		}

		datas := make([]string, cellcount)
		for i, cell := range row.Cells {

			if types[i] != "String" {
				if len(cell.String()) == 0 {
					datas[i] = "0"
				} else {
					datas[i] = cell.String()
				}
			} else {
				datas[i] = cell.String()
			}
		}
		csvwriter.Write(datas)
	}

	csvwriter.Flush()
}

func createProto(fileName string, packageName string, file *xlsx.File) {
	proto := "syntax = \"proto2\";\r\n"
	proto = proto + "package " + packageName + ";\r\n"
	proto = proto + "\r\n"

	for _, sheet := range file.Sheets {
		proto = proto + "message " + sheet.Name + " {\r\n"

		tableName := sheet.Rows[1].Cells[0].String()
		cellcount := len(sheet.Rows[6].Cells)

		types := make([]string, cellcount)
		for i, cell := range sheet.Rows[4].Cells {
			types[i] = cell.String()
		}
		proto = proto + "\tmessage " + tableName + " {\r\n"

		//columns := make([]string, cellcount)
		for i, cell := range sheet.Rows[6].Cells {
			proto = proto + "\t\trequired "
			if types[i] == "UInt32" {
				proto = proto + "uint32 "
			} else if types[i] == "String" {
				proto = proto + "string "
			}

			proto = proto + cell.String() + " = " + strconv.Itoa(i+1) + ";\r\n"
		}
		proto = proto + "\t}\r\n"
		proto = proto + "\trequired string TableName = 1;\r\n"
		proto = proto + "\trepeated " + tableName + " Data = 2;\r\n"

		proto = proto + "}\r\n\r\n"
	}

	outFile, err := os.Create(fmt.Sprintf("%s.proto", fileName))
	defer outFile.Close()
	if err != nil {
		return
	}

	outFile.WriteString(proto)
}
