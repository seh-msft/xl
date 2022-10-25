# xl

A tool for transliterating Excel `.xlsx` files into other, more programmatically useful, formats. 

Wraps [github.com/xuri/excelize/v2](https://godocs.io/github.com/xuri/excelize/v2). 

Written in [Go](https://golang.org).

## Build

	; go build

## Install

	; go install

## Usage

```
Usage of xl:
  -all
        Process all sheets
  -csv
        Output format should be CSV; implies Matrix mode
  -go
        Output format should be in Go syntax
  -i string
        Excel file to read from; default stdin
  -json
        Output format should be JSON
  -notitles
        Sheet does _not_ have column names as row 0; default has col names; forces Matrix mode
  -o string
        Output file to write to; default stdout
  -sheet string
        Excel sheet to search; empty uses first sheet in file
  -stats
        Print fun sheet statistics
  -striptitles
        Column names exist and should be elided from the output; forces Matrix mode
  -table
        Output should be a 2D matrix rather than a map→key object
```

## Examples

```
; xl -i .\dump.xlsx -json -o dump.json
info: #sheets read: 1 #cols: 10 #elements: 5010 #nrows: 501
;
```
