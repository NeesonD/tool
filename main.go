package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"sort"
)

var (
	filename = "filelist.txt"
	fromDir      = "Z:\\阿里云盘Open\\测试\\世界顶级畅销书530本"
	toDir = "./filelist.xlsx"
)

func main() {
	toExl()
}

func toExl() {
	// 新建一个Excel文件
	f := excelize.NewFile()
	// 新建一个Sheet，命名为"文件列表"
	sheetName := "文件列表"
	index, _ := f.NewSheet(sheetName)
	f.SetActiveSheet(index)
	f.SetCellValue(sheetName, fmt.Sprintf("A%d", 1), "序号")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", 1), "书名（按名称排序）")

	filenames := make([]string, 0)

	// 遍历目录下的文件
	rowIndex := 2
	err := filepath.Walk(fromDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		// 如果当前文件是文件夹，则跳过
		if info.IsDir() {
			return nil
		}
		filenames = append(filenames, info.Name())
		// 将文件名写入Sheet中

		return nil
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	sort.Strings(filenames)
	for i, s := range filenames {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowIndex), 1+i)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowIndex), s)
		rowIndex++
	}

	// 将文件保存到本地
	err = f.SaveAs(toDir)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("Done!")
}

func toTxt() {
	// 打开要写入的文件
	f, err := os.Create(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()

	// 遍历目录下的文件
	err = filepath.Walk(fromDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		// 如果当前文件是文件夹，则跳过
		if info.IsDir() {
			return nil
		}
		// 将文件名写入文件
		_, err = fmt.Fprintf(f, "%s\n", info.Name())
		if err != nil {
			return err
		}
		return nil
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("Done!")
}
