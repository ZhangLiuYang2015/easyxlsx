package easyxlsx

import (
	"github.com/tealeg/xlsx/v3"
)

func AnalysisByFile(file []byte, template SheetTemplate) (data []interface{}, err error) {
	excelFile, err := xlsx.OpenBinary(file)
	if err != nil {
		return data, err
	}

	sheets := excelFile.Sheets
	if len(sheets) == 0 {
		return
	}

	return analysisSheet(sheets[0], template)
}

func AnalysisByFilePath(path string, template SheetTemplate) (data []interface{}, err error) {
	excelFile, err := xlsx.OpenFile(path)
	if err != nil {
		return data, err
	}

	sheets := excelFile.Sheets
	if len(sheets) == 0 {
		return
	}

	return analysisSheet(sheets[0], template)
}
