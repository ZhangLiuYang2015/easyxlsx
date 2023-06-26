package easyxlsx

import "github.com/tealeg/xlsx/v3"

func analysisSheet(sheet *xlsx.Sheet, template SheetTemplate) (data []interface{}, err error) {
	template.transform()

	err = sheet.ForEachRow(func(r *xlsx.Row) error {
		num := r.GetCoordinate()
		if !template.isInvalid(num) {
			// nothing
			return nil
		}

		// handle row
		elem := template.newElem()
		err2 := r.ForEachCell(func(c *xlsx.Cell) error {
			// handle rank
			rank, _ := c.GetCoordinates()
			fieldName := template.fieldName(rank)
			if fieldName != "" {
				value := c.Value
				if c.Value == "" && len(c.RichText) != 0 {
					value = c.RichText[0].Text
				}
				elem.FieldByName(fieldName).SetString(value)
			}
			return nil
		})
		data = append(data, elem.Interface())
		return err2
	})
	return
}
