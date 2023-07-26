package main

import (
	"github.com/beevik/etree"
	"os"
	"path/filepath"
	"strings"
)

const docxTraceId = "rId9999"
const docxTraceType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
const docxTraceTemp = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId9999" Target="${traceUrl}" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"/>
</Relationships>`

func docxTraceInfo(traceUrl string) string {
	return strings.Replace(docxTraceTemp, "${traceUrl}", traceUrl, -1)
}

// GenerateTracerDocx 生成可追踪文档
func GenerateTracerDocx(srcFile, dstFile, traceUrl string) (err error) {
	var (
		tempDir  string
		document *etree.Document
	)

	// 1、解压 docx 文件
	tempDir, err = ExtractZip(srcFile, "file-trace-*")
	if err != nil {
		return err
	}

	// 2、添加/修改 settings.xml.rels 文件
	relsFile := filepath.Join(tempDir, "word", "_rels", "settings.xml.rels")
	if _, err = os.Stat(relsFile); err != nil {
		// 创建文件
		// 添加追踪信息
		err = os.WriteFile(relsFile, []byte(docxTraceInfo(traceUrl)), os.ModePerm)
		if err != nil {
			return err
		}
	} else {
		// 读取文件
		document, err = ReadXml(relsFile)
		if err != nil {
			return err
		}

		exist := false
		relationships := document.SelectElement("Relationships")
		for _, element := range relationships.ChildElements() {
			// 判断 Type 属性
			if element.SelectAttrValue("Type", "") == docxTraceType {
				// 判断 Target 属性是否为 traceUrl
				// 如果是，则标识为存在；如果不是，则替换为 traceUrl
				if element.SelectAttrValue("Target", "") == traceUrl {
					exist = true
					break
				} else {
					// 替换 traceUrl
					element.SelectAttr("Target").Value = traceUrl
				}
			}
		}

		// 存在文件，但是不存在追踪信息
		if !exist {
			// 添加节点
			node := relationships.CreateElement("Relationship")
			node.CreateAttr("Id", docxTraceId)
			node.CreateAttr("Type", docxTraceType)
			node.CreateAttr("Target", traceUrl)
			node.CreateAttr("TargetMode", "External")
			relationships.AddChild(node)
		}

		// 更新 settings.xml.rels 文件
		err = WriteXml(document, relsFile)
		if err != nil {
			return err
		}
	}

	// 3、修改 settings.xml 文件
	xmlFile := filepath.Join(tempDir, "word", "settings.xml")
	document, err = ReadXml(xmlFile)
	if err != nil {
		return err
	}

	exist := false
	settings := document.SelectElement("w:settings")
	for _, element := range settings.ChildElements() {
		// 判断是否存在 attachedTemplate 节点
		if element.Tag == "attachedTemplate" {
			if element.SelectAttr("r:id").Value == docxTraceId {
				exist = true
			} else {
				// 替换 traceId
				element.SelectAttr("r:id").Value = docxTraceId
			}
			break
		}
	}

	if !exist {
		// 添加节点
		node := settings.CreateElement("w:attachedTemplate")
		node.CreateAttr("r:id", docxTraceId)
		settings.InsertChildAt(1, node)
	}

	// 更新 settings.xml 文件
	err = WriteXml(document, xmlFile)
	if err != nil {
		return err
	}

	// 4、压缩文件夹，生成新的 docx 文件
	err = CompressZip(tempDir, dstFile)
	if err != nil {
		return err
	}

	// 5、删除临时目录
	err = os.RemoveAll(tempDir)
	if err != nil {
		return err
	}
	return nil
}
