package easyxlsx

import (
	"archive/zip"
	"encoding/csv"
	"github.com/tealeg/xlsx/v3"
	"io"
	"strings"
	"sync"
)

var (
	DefaultStyle = getDefaultStyle()
)

func getDefaultStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	font := *xlsx.NewFont(12, "Verdana")
	font.Bold = true
	style.Font = font
	return style
}

func writeSheet(file *xlsx.File, headerNames []string, data []interface{}) (err error) {
	sheet, _ := file.AddSheet("Sheet1")
	titleRow := sheet.AddRow()
	for i, name := range headerNames {
		cell := titleRow.AddCell()
		cell.Value = name
		cell.SetStyle(DefaultStyle)
		_ = sheet.SetColAutoWidth(i+1, func(s string) float64 {
			return float64(len(s) * 2)
		})
	}

	if len(data) != 0 {
		for i := range data {
			row := sheet.AddRow()
			row.WriteStruct(data[i], -1)

		}
	}

	return
}

func NewZipWriter(wr io.Writer, wLock *sync.Mutex) (zipWriter *ZipWriter) {
	zipWriter = new(ZipWriter)
	zipWriter.writer = zip.NewWriter(wr)
	zipWriter.wLock = wLock
	return
}

type ZipWriter struct {
	writer *zip.Writer
	wLock  *sync.Mutex
}

func (zipWriter *ZipWriter) Close() (err error) {
	return zipWriter.writer.Close()
}

func (zipWriter *ZipWriter) CompressCsv2Zip(csvName string, dataArr []interface{}) (err error) {
	// 1.将数据写入csv
	csvBuff, err := Write2Csv(dataArr)
	if err != nil {
		return
	}

	// 2.创建一个zip文件条目
	csvFile, err := zipWriter.writer.Create(csvName)
	if err != nil {
		return
	}

	// 3.将csv缓存数据写入文件条目
	_, err = csvFile.Write([]byte(csvBuff.String()))
	if err != nil {
		return
	}
	return nil
}

// ConcurrentCompress 顺序压缩文件
func (zipWriter *ZipWriter) ConcurrentCompress(csvName string, buff *strings.Builder) (err error) {
	// 1.是否存在写锁，存在则加锁，并在压缩完成后释放
	if zipWriter.wLock != nil {
		zipWriter.wLock.Lock()
	}
	defer func() {
		if zipWriter.wLock != nil {
			zipWriter.wLock.Unlock()
		}
		if err2 := recover(); err2 != nil {
		}
	}()

	// 2.创建一个zip文件条目
	csvFile, err := zipWriter.writer.Create(csvName)
	if err != nil {
		return
	}

	// 3.将缓存数据写入文件条目
	_, err = csvFile.Write([]byte(buff.String()))
	return nil
}

func Write2Csv(dataArr []interface{}) (buff *strings.Builder, err error) {
	buff = new(strings.Builder)
	writer := csv.NewWriter(buff)

	stringArr := Convert2StringArr(dataArr)
	err = writer.WriteAll(stringArr)
	writer.Flush()
	return
}
