package main

import (
	_ "embed"
	"fmt"
	"github.com/beevik/etree"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

//go:embed assets/ninelock.png
var markImage []byte

//go:embed assets/drawing.xml
var xlsxDrawing []byte

//go:embed assets/drawing.xml.tpl
var xlsxDrawingTpl string

const xlsxTraceId = "rId9999"
const xlsxTraceType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
const xlsxTraceTemp = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Target="/xl/media/image1.png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
    <Relationship Id="rId9999" Target="${traceUrl}" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
</Relationships>`

const xlsxSheetRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Target="/xl/drawings/drawing1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
</Relationships>`

func xlsxTraceInfo(traceUrl string) string {
	return strings.Replace(xlsxTraceTemp, "${traceUrl}", traceUrl, -1)
}

// GenerateTracerXlsx 生成可追踪表格
func GenerateTracerXlsx(srcFile, dstFile, traceUrl string) (err error) {
	var (
		tempDir         string
		mediaDir        string
		sheetDir        string
		sheetRelsDir    string
		drawingsDir     string
		drawingsRelsDir string

		count         int
		xlsxImageId   = "rId1"
		xlsxImageName = "image1.png"

		document   *etree.Document
		xmlContent []byte
	)

	// 1、解压 xlsx 文件
	tempDir, err = ExtractZip(srcFile, "file-trace-*")
	if err != nil {
		return err
	}

	// 2、添加图片
	mediaDir = filepath.Join(tempDir, "xl", "media")
	err = CreateDir(mediaDir)
	if err != nil {
		return err
	}

	count, err = SubFileCount(mediaDir)
	if err != nil {
		return err
	}

	xlsxImageId = fmt.Sprintf("rId%d", count+1)
	xlsxImageName = fmt.Sprintf("image%d.png", count+1)
	err = os.WriteFile(filepath.Join(mediaDir, xlsxImageName), markImage, os.ModePerm)
	if err != nil {
		return err
	}

	// 3、添加/修改 drawing1.xml，drawing1.xml.rels 文件
	drawingsDir = filepath.Join(tempDir, "xl", "drawings")
	drawingsRelsDir = filepath.Join(tempDir, "xl", "drawings", "_rels")
	err = CreateDir(drawingsRelsDir)
	if err != nil {
		return err
	}

	xmlFile := filepath.Join(drawingsDir, "drawing1.xml")
	relsFile := filepath.Join(drawingsRelsDir, "drawing1.xml.rels")
	if _, err = os.Stat(relsFile); err != nil {
		count = 0 // 认为不存在
		// 创建文件
		// 添加追踪信息
		err = os.WriteFile(relsFile, []byte(xlsxTraceInfo(traceUrl)), os.ModePerm)
		if err != nil {
			return err
		}

		// 添加 png 到 [Content_Types].xml 文件
		typesFile := filepath.Join(tempDir, "[Content_Types].xml")
		document, err = ReadXml(typesFile)
		if err != nil {
			return err
		}

		types := document.SelectElement("Types")
		node := types.CreateElement("Default")
		node.CreateAttr("Extension", "png")
		node.CreateAttr("ContentType", "image/png")
		types.AddChild(node)

		// 更新 [Content_Types].xml 文件
		err = WriteXml(document, typesFile)
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
			// 判断 Id 属性
			if element.SelectAttrValue("Id", "") == xlsxTraceId {
				// 判断 Target 属性是否为 traceUrl
				// 如果是，则标识为存在；如果不是，则替换为 traceUrl
				if element.SelectAttrValue("Target", "") == traceUrl {
					exist = true
					break
				} else {
					element.CreateAttr("Target", traceUrl)
					exist = true
					break
				}
			}
		}

		// 存在文件，但是不存在追踪信息
		if !exist {
			// 添加图片节点
			node := relationships.CreateElement("Relationship")
			node.CreateAttr("Id", xlsxImageId)
			node.CreateAttr("Type", xlsxTraceType)
			node.CreateAttr("Target", fmt.Sprintf("/xl/media/%s", xlsxImageName))
			relationships.AddChild(node)

			// 添加追踪节点
			node = relationships.CreateElement("Relationship")
			node.CreateAttr("Id", xlsxTraceId)
			node.CreateAttr("Type", xlsxTraceType)
			node.CreateAttr("Target", traceUrl)
			node.CreateAttr("TargetMode", "External")
			relationships.AddChild(node)
		}

		// 更新 drawing1.xml.rels 文件
		err = WriteXml(document, relsFile)
		if err != nil {
			return err
		}
	}

	if _, err = os.Stat(xmlFile); err != nil {
		if count == 0 {
			// 创建文件
			// 添加追踪信息
			err = os.WriteFile(xmlFile, xlsxDrawing, os.ModePerm)
			if err != nil {
				return err
			}

			// 添加 drawing1.xml 到 [Content_Types].xml 文件
			typesFile := filepath.Join(tempDir, "[Content_Types].xml")
			document, err = ReadXml(typesFile)
			if err != nil {
				return err
			}

			types := document.SelectElement("Types")
			node := types.CreateElement("Override")
			node.CreateAttr("PartName", "/xl/drawings/drawing1.xml")
			node.CreateAttr("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
			types.AddChild(node)

			// 更新 [Content_Types].xml 文件
			err = WriteXml(document, typesFile)
			if err != nil {
				return err
			}
		} else {
			// 不考虑这个情况
		}
	} else {
		// 读取文件内容
		xmlContent, err = os.ReadFile(xmlFile)
		if err != nil {
			return err
		}

		if !strings.Contains(string(xmlContent), xlsxTraceId) {
			xlsxDrawingTpl = strings.Replace(xlsxDrawingTpl, "${no}", strconv.Itoa(count+1), -1)
			xlsxDrawingTpl = ">" + xlsxDrawingTpl + "</xdr:wsDr>"

			current := strings.Replace(string(xmlContent), "></xdr:wsDr>", xlsxDrawingTpl, -1)

			err = os.WriteFile(xmlFile, []byte(current), os.ModePerm)
			if err != nil {
				return err
			}

			// 添加 r 的命名空间
			document, err = ReadXml(xmlFile)
			if err != nil {
				return err
			}

			root := document.SelectElement("xdr:wsDr")
			root.CreateAttr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

			err = WriteXml(document, xmlFile)
			if err != nil {
				return err
			}
		}
	}

	// 4、添加/修改 sheet1.xml，sheet1.xml.rels 文件
	sheetDir = filepath.Join(tempDir, "xl", "worksheets")
	sheetRelsDir = filepath.Join(tempDir, "xl", "worksheets", "_rels")
	err = CreateDir(sheetRelsDir)
	if err != nil {
		return err
	}

	xmlFile = filepath.Join(sheetDir, "sheet1.xml")
	relsFile = filepath.Join(sheetRelsDir, "sheet1.xml.rels")
	if _, err = os.Stat(relsFile); err != nil {
		// 创建文件
		// 写入 drawing 信息
		err = os.WriteFile(relsFile, []byte(xlsxSheetRels), os.ModePerm)
		if err != nil {
			return err
		}
	}

	if _, err = os.Stat(xmlFile); err == nil {
		document, err = ReadXml(xmlFile)
		if err != nil {
			return err
		}

		exist := false
		workSheet := document.SelectElement("worksheet")
		for _, element := range workSheet.ChildElements() {
			// 判断是否存在 drawing 节点
			if element.Tag == "drawing" {
				exist = true
				break
			}
		}

		if !exist {
			// 添加 drawing 节点
			node := workSheet.CreateElement("drawing")
			node.CreateAttr("r:id", "rId1")
			node.CreateAttr("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
			workSheet.AddChild(node)

			// 更新 sheet1.xml 文件
			err = WriteXml(document, xmlFile)
			if err != nil {
				return err
			}
		}
	}

	// 5、压缩文件夹，生成新的 xlsx 文件
	err = CompressZip(tempDir, dstFile)
	if err != nil {
		return err
	}

	// 6、删除临时目录
	err = os.RemoveAll(tempDir)
	if err != nil {
		return err
	}
	return nil
}
