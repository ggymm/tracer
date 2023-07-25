package main

import (
	"github.com/beevik/etree"
	"os"
	"path/filepath"
	"strings"
)

var traceId = "rId1"
var traceInfo = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
	Target="${traceUrl}"
	TargetMode="External"/>
</Relationships>`

func GenerateTracerDocx(srcFile, dstFile, traceUrl string) (err error) {
	var (
		tempDir string

		exist    = false
		document *etree.Document
	)

	tempDir, err = ExtractZip(srcFile, "file-trace-*")
	if err != nil {
		return nil
	}

	// 判断 word/_rels/ 目录下是否包含 settings.xml.rels 文件
	relsFile := filepath.Join(tempDir, "word", "_rels", "settings.xml.rels")
	if _, err = os.Stat(relsFile); err != nil {
		// 替换 traceUrl
		traceInfo = strings.Replace(traceInfo, "${traceUrl}", traceUrl, -1)

		// 写入 settings.xml.rels
		err = os.WriteFile(relsFile, []byte(traceInfo), os.ModePerm)
		if err != nil {
			return err
		}
	} else {
		// 修改
		traceId = "rId2"
	}

	// 判断 word/settings.xml 文件是否包含 w:attachedTemplate 标签
	settings := filepath.Join(tempDir, "word", "settings.xml")
	document, err = ReadXml(settings)
	if err != nil {
		return err
	}

	root := document.SelectElement("w:settings")
	for _, element := range root.ChildElements() {
		if element.Space == "w" && element.Tag == "attachedTemplate" {
			exist = true
			break
		}
	}

	if !exist {
		// 添加节点
		node := root.CreateElement("w:attachedTemplate")
		node.CreateAttr("r:id", traceId)

		root.InsertChildAt(1, node)
	}

	err = WriteXml(document, settings)
	if err != nil {
		return err
	}

	// 压缩文件夹，生成新的 docx 文件
	err = CompressDir(tempDir, dstFile)
	if err != nil {
		return err
	}

	// 删除临时目录
	err = os.RemoveAll(tempDir)
	if err != nil {
		return err
	}

	return nil
}
