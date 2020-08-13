package main

import (
	"flag"
	"fmt"
	"strings"

	"github.com/tealeg/xlsx"
)

// Row is a line of CSV
type Row struct {
	Name     string
	Type     string
	Quantity string
	Maker    string
}

func (r Row) String() string {
	ss := []string{r.Name, r.Type, r.Quantity, r.Maker}
	return fmt.Sprintf("%s", strings.Join(ss, ","))
}

func main() {
	flag.Parse()
	for _, file := range flag.Args() {
		excel, err := xlsx.OpenFile(file)
		if err != nil {
			fmt.Printf(err.Error())
		}

		sheet1 := excel.Sheets[0]

		r := Row{}
		r.Name = strings.ReplaceAll(sheet1.Rows[10].Cells[2].Value, "\n", "")
		r.Type = strings.ReplaceAll(sheet1.Rows[10].Cells[4].Value, "\n", "")
		r.Quantity = sheet1.Rows[10].Cells[6].Value
		r.Maker = strings.ReplaceAll(sheet1.Rows[10].Cells[10].Value, "\n", "")
		fmt.Println(r)
	}
}
