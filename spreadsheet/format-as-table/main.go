// Copyright 2026 FoxyUtils ehf. All rights reserved.
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/unidoc/unioffice/v2/common/license"
	"github.com/unidoc/unioffice/v2/schema/soo/sml"
	"github.com/unidoc/unioffice/v2/spreadsheet"
)

func init() {
	// Make sure to load your metered License API key prior to using the library.
	// If you need a key, you can sign up and create a free one at https://cloud.unidoc.io
	err := license.SetMeteredKey(os.Getenv(`UNIDOC_LICENSE_API_KEY`))
	if err != nil {
		panic(err)
	}
}

func main() {
	ss := spreadsheet.New()
	defer ss.Close()
	// add a single sheet
	sheet := ss.AddSheet()

	// header row - the table derives its column names from this row
	hdrRow := sheet.AddRow()
	hdrRow.AddCell().SetString("Region")
	hdrRow.AddCell().SetString("Quarter")
	hdrRow.AddCell().SetString("Sales")

	// data rows
	data := []struct {
		region  string
		quarter string
		sales   float64
	}{
		{"North", "Q1", 1500},
		{"North", "Q2", 1800},
		{"South", "Q1", 1200},
		{"South", "Q2", 1650},
		{"West", "Q1", 2100},
		{"West", "Q2", 1950},
	}
	for _, d := range data {
		row := sheet.AddRow()
		row.AddCell().SetString(d.region)
		row.AddCell().SetString(d.quarter)
		row.AddCell().SetNumber(d.sales)
	}

	// Derive all ranges from the data length so they stay correct if the data
	// above changes. The header is row 1, data occupies rows 2..lastDataRow,
	// and the totals row sits just below at totalsRow.
	const firstDataRow = 2
	lastDataRow := firstDataRow + len(data) - 1
	totalsRow := lastDataRow + 1
	dataRange := fmt.Sprintf("A1:C%d", lastDataRow)
	withTotalsRange := fmt.Sprintf("A1:C%d", totalsRow)
	salesRange := fmt.Sprintf("C%d:C%d", firstDataRow, lastDataRow)

	// Convert the range to a "Format as Table" range. The range must cover the
	// header row plus at least one data row. This applies an AutoFilter on the
	// header row and a banded (alternating row color) style by default.
	tbl := sheet.AddTable(dataRange, "SalesTable")

	// Pick a built-in style and emphasize the first column.
	tbl.SetStyle(spreadsheet.TableStyleMedium9)
	tbl.SetShowFirstColumn(true)

	// Add a totals row that sums the Sales column. The totals row occupies the
	// last row of the table reference, so extend the table by one row first.
	tbl.SetReference(withTotalsRange)
	tbl.SetTotalsRow(true)
	if firstCol, ok := tbl.Column(0); ok {
		firstCol.SetTotalsRowLabel("Total")
	}
	if salesCol, ok := tbl.Column(2); ok {
		salesCol.SetTotalsRowFunction(sml.ST_TotalsRowFunctionSum)
	}

	// Materialize the totals row cells so the values show when opened.
	row := sheet.AddRow()
	row.AddCell().SetString("Total")
	row.AddCell() // empty cell under Quarter
	// SUBTOTAL with function 109 sums the data rows, ignoring filtered rows.
	row.AddCell().SetFormulaRaw(fmt.Sprintf("SUBTOTAL(109,%s)", salesRange))

	if err := ss.Validate(); err != nil {
		log.Fatalf("error validating sheet: %s", err)
	}

	if err := ss.SaveToFile("format-as-table.xlsx"); err != nil {
		log.Fatalf("error saving spreadsheet: %s", err)
	}
}
