// Copyright (c) 2022, Microsoft Corporation, Sean Hinchee
// Licensed under the MIT License.

package main

import (
	"bufio"
	"encoding/csv"
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"strings"

	xl "github.com/xuri/excelize/v2"
)

type Mode int

const (
	Map Mode = iota
	MultiSheet
	Matrix
	Stats
)

var (
	allSheets     = flag.Bool("all", false, "Process all sheets")
	useSheet      = flag.String("sheet", "", "Excel sheet to search; empty uses first sheet in file")
	noColNames    = flag.Bool("notitles", false, "Sheet does _not_ have column names as row 0; default has col names; forces Matrix mode")
	stripColNames = flag.Bool("striptitles", false, "Column names exist and should be elided from the output; forces Matrix mode")
	tableMode     = flag.Bool("table", false, "Output should be a 2D matrix rather than a map→key object")
	statsMode     = flag.Bool("stats", false, "Print fun sheet statistics")
	asJson        = flag.Bool("json", false, "Output format should be JSON")
	asGo          = flag.Bool("go", false, "Output format should be in Go syntax")
	asCSV         = flag.Bool("csv", false, "Output format should be CSV; implies Matrix mode")
	//useAlphaTitles = flag.Bool("alphatitles", false, "Rather than using col[0] as the title, use the convention A0, B0, etc.")

	inPath  = flag.String("i", "", "Excel file to read from; default stdin")
	outPath = flag.String("o", "", "Output file to write to; default stdout")
)

func main() {
	mode := Map                                     // Used in Matrix mode
	bookTab := make(map[string]map[string][]string) // If using all sheets and table format per-sheet
	bookMat := make(map[string][][]string)          // If using all sheets 2D matrix format per-sheet

	in := bufio.NewReader(os.Stdin)
	out := bufio.NewWriter(os.Stdout)

	flag.Parse()

	if *tableMode || *stripColNames || *asCSV {
		mode = Matrix
	}
	if *statsMode {
		mode = Stats
	}
	if !*asJson && !*asGo && !*asCSV {
		mode = Stats
	}

	if *inPath != "" {
		f, err := os.Open(*inPath)
		efatal(err, "could not open input file")
		defer f.Close()
		in = bufio.NewReader(f)
	}

	if *outPath != "" {
		f, err := os.Create(*outPath)
		efatal(err, "could not create output file")
		defer f.Close()
		out = bufio.NewWriter(f)
	}

	defer out.Flush()

	opts := xl.Options{}
	xf, err := xl.OpenReader(in, opts)
	efatal(err, "could not read input excel")
	defer xf.Close()

	sheets := xf.GetSheetList()
	nSheets := 0
	nRows := 0
	nCols := 0
	sheetFound := false
	rowSize := 0

	for _, sheet := range sheets {
		if *useSheet != "" && sheet != *useSheet {
			continue
		}
		sheetFound = true
		bookTab[sheet] = make(map[string][]string)
		bookMat[sheet] = [][]string{}
		nSheets++
		cols, err := xf.Cols(sheet)
		efatal(err, "could not get columns for sheet", sheet)

		for cols.Next() {
			nCols++
			col, err := cols.Rows()
			// Might be erroneous for titled/nontitled mode
			rowSize = len(col)
			efatal(err, "could not get rows of col for sheet", sheet)

			switch mode {
			case Map:
				// Assumes we have a title
				if len(col) < 1 {
					// Column with NO title and NO values
					fatal("can't use Map mode with no title or values; col #:", nCols-1, "sheet:", sheet)
				} else if len(col) < 2 {
					// Column with title and NO values (probably)
					bookTab[sheet][col[0]] = []string{}
				} else {
					// Column has title and values
					bookTab[sheet][col[0]] = col[1:]
				}
			case Matrix:
				// Table format across all sheets
				bookMat[sheet] = append(bookMat[sheet], col)
			default:
				// Stats mode does nothing
			}

			for rowi, rowCell := range col {
				if !*noColNames && rowi == 0 && len(strings.TrimSpace(rowCell)) > 0 {
					if mode == Stats {
						fmt.Fprintln(out, "Column name:", `"`+rowCell+`"`, "at col#", nCols-1, "with", len(col), "rows")
					}
				}
				nRows++
			}
		}

		if !*allSheets {
			break
		}
	}

	fmt.Fprintln(os.Stderr, "info: #sheets read:", nSheets, "#cols:", nCols, "#elements:", nRows, "#nrows:", rowSize)

	if !sheetFound {
		fatal("could not find sheet by name of:", *useSheet)
	}

	// JSON mode
	if *asJson {
		enc := json.NewEncoder(out)
		switch mode {
		case Matrix:
			efatal(enc.Encode(bookMat), "could not JSON encode")
		case Map:
			efatal(enc.Encode(bookTab), "could not JSON encode")
		}

		return
	}

	// Go syntax mode
	if *asGo {
		switch mode {
		case Matrix:
			fmt.Fprintf(out, "%#v\n", bookMat)
		case Map:
			fmt.Fprintf(out, "%#v\n", bookTab)
		}

		return
	}

	// CSV mode
	if *asCSV {
		// Implicitly matrix mode
		w := csv.NewWriter(out)
		defaultSheet := sheets[0]
		// fmt.Println(defaultSheet)
		var tab [][]string
		records := bookMat[defaultSheet]
		nCols := len(records)
		var nRows int = 0
		for ci := 0; ci < nCols; ci++ {
			if len(records[ci]) > nRows {
				nRows = len(records[ci])
			}
		}

		tab = make([][]string, nRows)
		for i := 0; i < len(tab); i++ {
			tab[i] = make([]string, nCols)
		}

		for ci := 0; ci < len(records); ci++ {
			for ri := 0; ri < len(records[ci]); ri++ {
				if *noColNames && ri == 0 {
					continue
				}
				tab[ri][ci] = records[ci][ri]
			}
		}
		err := w.WriteAll(tab)
		efatal(err, "could not write output CSV")

		return
	}
}

func efatal(err error, s ...any) {
	if err == nil {
		return
	}
	var msg []any = []any{"err:"}
	msg = append(msg, s...)
	msg = append(msg, "->", err.Error())
	fatal(msg...)
}

func fatal(s ...any) {
	fmt.Fprintln(os.Stderr, s...)
	os.Exit(1)
}
