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

		r := Row{}
		for _, sheet := range excel.Sheets {
			for _, row := range sheet.Rows {
				r.Name = strings.ReplaceAll(row.Cells[2].Value, "\n", "")
				r.Type = strings.ReplaceAll(row.Cells[4].Value, "\n", "")
				r.Quantity = row.Cells[6].Value
				r.Maker = strings.ReplaceAll(row.Cells[10].Value, "\n", "")
				fmt.Println(r)
			}
		}
	}
}
