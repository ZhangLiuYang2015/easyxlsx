package easyxlsx

import (
	"reflect"
	"strings"
)

const (
	Name = "name"
)

type Handler interface {
	Handle([]interface{})
}

type SheetTemplate struct {
	Template    interface{}
	mapping     map[int]string
	HeaderNames []string
	pType       reflect.Type

	// valid data rows
	Begin int
	End   int

	SheetName string

	// 解析后处理器
	Handler Handler
}

func (st *SheetTemplate) transform() {
	st.pType = reflect.TypeOf(st.Template)
	st.mapping = make(map[int]string)
	for i := 0; i < st.pType.NumField(); i++ {
		field := st.pType.Field(i)
		xlsx := field.Tag.Get("xlsx")
		st.mapping[i] = field.Name
		if xlsx != "" {
			split := strings.Split(xlsx, ";")
			for _, tag := range split {
				split2 := strings.Split(tag, ":")
				if split2[0] == Name {
					st.HeaderNames = append(st.HeaderNames, split2[1])
				}
			}

		}
	}

	return
}

func (st *SheetTemplate) getHeaderNames() []string {
	return st.HeaderNames
}

func (st *SheetTemplate) newElem() reflect.Value {
	return reflect.New(st.pType).Elem()
}

func (st *SheetTemplate) isInvalid(curRow int) bool {
	if curRow == 0 {
		return false
	}

	if st.Begin == 0 && st.End == st.Begin {
		return true
	}

	return curRow >= st.Begin && curRow <= st.End
}

func (st *SheetTemplate) fieldName(rank int) string {
	return st.mapping[rank]
}

func Convert2StringArr(dataArr []interface{}) (arr [][]string) {
	if len(dataArr) == 0 {
		return
	}

	// 解析标题行
	data := dataArr[0]
	pType := reflect.TypeOf(data).Elem()
	var headerNames, fieldNames []string
	for i := 0; i < pType.NumField(); i++ {
		field := pType.Field(i)
		fieldNames = append(fieldNames, field.Name)
		xlsx := field.Tag.Get("xlsx")
		if xlsx != "" {
			split := strings.Split(xlsx, ";")
			for _, tag := range split {
				split2 := strings.Split(tag, ":")
				if split2[0] == Name {
					headerNames = append(headerNames, split2[1])
				}
			}
		}
	}
	arr = append(arr, headerNames)

	// 解析数据
	for _, d := range dataArr {
		vl := reflect.ValueOf(d).Elem()
		var vlArr []string
		for _, name := range fieldNames {
			byName := vl.FieldByName(name)
			vlArr = append(vlArr, byName.String())
		}
		arr = append(arr, vlArr)
	}
	return
}
