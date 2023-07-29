package utils

import (
	"github.com/beevik/etree"
	"os"
)

// ReadXml 读取 xml 文件
// filename: 文件名
func ReadXml(filename string) (document *etree.Document, err error) {
	document = etree.NewDocument()
	err = document.ReadFromFile(filename)
	if err != nil {
		return nil, err
	}

	return document, nil
}

// WriteXml 输出 xml 文件
// doc: etree.Document 对象
// filename: 文件名
func WriteXml(document *etree.Document, filename string) error {
	// 如果文件存在，则删除
	if _, err := os.Stat(filename); err == nil {
		err = os.Remove(filename)
		if err != nil {
			return err
		}
	}
	return document.WriteToFile(filename)
}
