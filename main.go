package main

import (
	"flag"
	"fmt"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	// VERSION plst version
	VERSION = "v0.1.0"
)

var (
	showVersion bool
	showHelp    bool
)

// Row is a line of CSV
type Row struct {
	Name     string
	Model    string
	Quantity string
	Maker    string
}

func (r Row) String() string {
	ss := []string{r.Name, r.Model, r.Quantity, r.Maker}
	return fmt.Sprintf("%s", strings.Join(ss, ","))
}

func main() {
	flag.BoolVar(&showVersion, "v", false, "show version")
	flag.BoolVar(&showVersion, "version", false, "show version")
	flag.BoolVar(&showHelp, "h", false, "show help")
	flag.BoolVar(&showHelp, "help", false, "show help")
	flag.Parse()

	// Show Version
	if showVersion {
		fmt.Println("plst version:", VERSION)
		return // versionを表示して終了
	}

	// Show Help
	if showHelp {
		fmt.Println(`
部品諸元表( .xlsx形式 )から、[ 品名, 型式, 数量, メーカ ] を抜き出します。

Usage:
	# 品名、型式、数量、メーカをSK553808.csvに保存
	$ plst SK553808_同調増幅回路_VLFANT01_部品諸元表.xlsx > SK553808.csv

	# 品名、型式、数量、メーカを標準出力に表示
	$ plst SK553808_同調増幅回路_VLFANT01_部品諸元表.xlsx | column -s, -t

	# 複数の部品諸元表から品名、型式、数量、メーカを標準出力に表示
	$ plst *.xlsx | column -s, -t
		`)
		return // helpを表示して終了
	}

	// Main
	for _, file := range flag.Args() {
		excel, err := xlsx.OpenFile(file)
		if err != nil {
			fmt.Printf(err.Error())
		}

		r := Row{}
		skip := true
		for _, sheet := range excel.Sheets {
			for _, row := range sheet.Rows {
				q := r // ひとつ前の行("同上"を処理するために必要)

				name := strings.ReplaceAll(row.Cells[2].Value, "\n", "")
				if strings.Join(strings.Fields(name), "") == "品名" { // "品名"がくるまでは出力しない
					skip = false
					continue
				}
				if name == "" || skip { // 品名が空なら次の行へ
					continue
				}

				if strings.HasPrefix(strings.Join(strings.Fields(name), ""), "欠番") { // 欠番で始まるとき
					continue
				}

				if strings.HasPrefix(strings.Join(strings.Fields(name), ""), "同上") { // 同上で始まるとき
					r.Name = q.Name // 前の行の値を適用
				} else {
					r.Name = name // 同上でなければ読み込んだ値を適用
				}

				model := strings.ReplaceAll(row.Cells[4].Value, "\n", "")
				if strings.HasPrefix(strings.Join(strings.Fields(model), ""), "同上") {
					r.Model = q.Model // 前の行の値を適用
				} else {
					r.Model = model // 同上でなければ読み込んだ値を適用
				}

				r.Quantity = row.Cells[6].Value

				maker := strings.ReplaceAll(row.Cells[10].Value, "\n", "")
				if strings.HasPrefix(strings.Join(strings.Fields(maker), ""), "同上") {
					r.Maker = q.Maker // 前の行の値を適用
				} else {
					r.Maker = maker // 同上でなければ読み込んだ値を適用
				}

				fmt.Println(r)
			}
		}
	}
}
