package main

import (
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
	excel, err := xlsx.OpenFile("./SK553808_同調増幅回路_VLFANT01_部品諸元表.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}

	sheet1 := excel.Sheets[0]
	// fmt.Println(sheet1.Name)

	r := Row{}
	r.Name = strings.ReplaceAll(sheet1.Rows[10].Cells[2].Value, "\n", "")
	r.Type = strings.ReplaceAll(sheet1.Rows[10].Cells[4].Value, "\n", "")
	// q := sheet1.Rows[10].Cells[6].Value
	// r.Quantity, err = strconv.Atoi(q)
	r.Quantity = sheet1.Rows[10].Cells[6].Value
	r.Maker = strings.ReplaceAll(sheet1.Rows[10].Cells[10].Value, "\n", "")
	// if err != nil {
	// 	fmt.Printf(err.Error())
	// }

	// fmt.Printf("%#v", r)
	fmt.Println(r)
}
