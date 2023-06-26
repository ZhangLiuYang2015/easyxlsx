package easyxlsx

import (
	"reflect"
	"strconv"
	"strings"
)

const (
	column = "column"
)

type SheetTemplate struct {
	Template interface{}
	mapping  map[int]string
	pType    reflect.Type

	// valid data rows
	Begin int
	End   int

	SheetName string
}

func (st *SheetTemplate) transform() {
	st.pType = reflect.TypeOf(st.Template)
	st.mapping = make(map[int]string)
	for i := 0; i < st.pType.NumField(); i++ {
		field := st.pType.Field(i)
		xlsx := field.Tag.Get("xlsx")
		if xlsx != "" {
			split := strings.Split(xlsx, ";")
			for _, tag := range split {
				split2 := strings.Split(tag, ":")
				if split2[0] == column {
					i, err := strconv.Atoi(split2[1])
					if err == nil {
						st.mapping[i] = field.Name
					}
				}
			}
		}
	}

	return
}

func (st *SheetTemplate) newElem() reflect.Value {
	return reflect.New(st.pType).Elem()
}

func (st *SheetTemplate) isInvalid(curRow int) bool {
	if st.Begin == 0 && st.End == st.Begin {
		return true
	}

	return curRow >= st.Begin && curRow <= st.End
}

func (st *SheetTemplate) fieldName(rank int) string {
	return st.mapping[rank]
}
