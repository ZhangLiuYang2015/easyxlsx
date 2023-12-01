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

	data, err = analysisSheet(sheets[0], template)
	if err != nil || len(data) == 0 || template.Handler == nil {
		return
	}

	template.Handler.Handle(data)
	return data, err
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

func AnalysisSheetsByFile2(file []byte, template SheetTemplate) (data map[string][]interface{}) {
	data = make(map[string][]interface{})
	excelFile, err := xlsx.OpenBinary(file)
	if err != nil {
		return
	}

	sheets := excelFile.Sheets
	if len(sheets) == 0 {
		return
	}

	for _, sheet := range sheets {
		rs, err := analysisSheet(sheet, template)
		if err == nil && len(rs) != 0 {
			data[sheet.Name] = rs
		}
	}
	return
}

func Export(template SheetTemplate, data []interface{}) (file *xlsx.File, err error) {
	file = xlsx.NewFile()

	template.transform()
	names := template.getHeaderNames()
	if len(names) == 0 {
		return
	}

	err = writeSheet(file, names, data)
	return
}
