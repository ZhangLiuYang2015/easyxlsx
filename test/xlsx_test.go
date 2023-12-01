package test

import (
	"fmt"
	"github.com/ZhangLiuYang2015/easyxlsx/easyxlsx"
	"testing"
)

type Obj struct {
	Ip     string `xlsx:"column:0"`
	Number string `xlsx:"column:1"`
}

func Test(t *testing.T) {
	template := easyxlsx.SheetTemplate{
		Template: Obj{},
	}

	data, err := easyxlsx.AnalysisByFilePath("C:\\Users\\zhangliuyang\\Desktop\\液冷IPMI&IP-20230529.xlsx", template)
	if err == nil {
		fmt.Println(data)
	}
}
