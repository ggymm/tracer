package ms_office

import (
	_ "embed"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"tracer/internal/assets"
	"tracer/pkg/utils"

	"github.com/beevik/etree"
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

// GenTracerDOCX 生成可追踪文档
func GenTracerDOCX(srcFile, dstFile, traceUrl string) (err error) {
	var (
		tempDir  string
		document *etree.Document
	)

	// 1、解压 docx 文件
	tempDir, err = utils.ExtractZip(srcFile, "file-trace-*")
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
		document, err = utils.ReadXml(relsFile)
		if err != nil {
			return err
		}

		exist := false
		relationships := document.SelectElement("Relationships")
		for _, element := range relationships.ChildElements() {
			// 判断 Id 属性
			if element.SelectAttrValue("Id", "") == docxTraceId {
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
		err = utils.WriteXml(document, relsFile)
		if err != nil {
			return err
		}
	}

	// 3、修改 settings.xml 文件
	xmlFile := filepath.Join(tempDir, "word", "settings.xml")
	document, err = utils.ReadXml(xmlFile)
	if err != nil {
		return err
	}

	exist := false
	settings := document.SelectElement("w:settings")
	for _, element := range settings.ChildElements() {
		// 判断是否存在 attachedTemplate 节点
		if element.Space == "w" && element.Tag == "attachedTemplate" {
			exist = true
			if element.SelectAttr("r:id").Value != docxTraceId {
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
	err = utils.WriteXml(document, xmlFile)
	if err != nil {
		return err
	}

	// 4、压缩文件夹，生成新的 docx 文件
	err = utils.CompressZip(tempDir, dstFile)
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

const pptxTraceId = "rId9999"
const pptxTraceType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
const pptxTraceTemp = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId9999" Target="${traceUrl}" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" />
</Relationships>`

func pptxTraceInfo(traceUrl string) string {
	return strings.Replace(pptxTraceTemp, "${traceUrl}", traceUrl, -1)
}

// GenTracerPPTX 生成可追踪演示文稿
func GenTracerPPTX(srcFile, dstFile, traceUrl string) (err error) {
	var (
		tempDir  string
		document *etree.Document
	)

	// 1、解压 pptx 文件
	tempDir, err = utils.ExtractZip(srcFile, "file-trace-*")
	if err != nil {
		return err
	}

	// 2、添加/修改 slide1.xml.rels 文件
	relsFile := filepath.Join(tempDir, "ppt", "slides", "_rels", "slide1.xml.rels")
	if _, err = os.Stat(relsFile); err != nil {
		// 创建文件
		// 添加追踪信息
		err = os.WriteFile(relsFile, []byte(pptxTraceInfo(traceUrl)), os.ModePerm)
		if err != nil {
			return err
		}
	} else {
		// 读取文件
		document, err = utils.ReadXml(relsFile)
		if err != nil {
			return err
		}

		exist := false
		relationships := document.SelectElement("Relationships")
		for _, element := range relationships.ChildElements() {
			// 判断 Id 属性
			if element.SelectAttrValue("Id", "") == pptxTraceId {
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
			node.CreateAttr("Id", pptxTraceId)
			node.CreateAttr("Type", pptxTraceType)
			node.CreateAttr("Target", traceUrl)
			node.CreateAttr("TargetMode", "External")
			relationships.AddChild(node)
		}

		// 更新 slide1.xml.rels 文件
		err = utils.WriteXml(document, relsFile)
		if err != nil {
			return err
		}
	}

	// 3、修改 slide1.xml 文件
	xmlFile := filepath.Join(tempDir, "ppt", "slides", "slide1.xml")
	document, err = utils.ReadXml(xmlFile)
	if err != nil {
		return err
	}

	exist := false
	slides := document.FindElement("p:sld > p:cSld > p:spTree > p:pic > p:blipFill > a:blip")
	if slides != nil {
		exist = true
		if slides.SelectAttr("r:link").Value != pptxTraceId {
			// 替换 traceId
			slides.SelectAttr("r:link").Value = pptxTraceId
		}
	}

	if !exist {
		tree := document.FindElement("//p:sld/p:cSld/p:spTree")
		nodeId := strconv.Itoa(len(tree.ChildElements()) + 1)
		assets.MSSlideTpl = strings.Replace(assets.MSSlideTpl, "${id}", nodeId, -1)
		assets.MSSlideTpl = strings.Replace(assets.MSSlideTpl, "${pptxTraceId}", pptxTraceId, -1)

		n := etree.NewDocument()
		err = n.ReadFromString(assets.MSSlideTpl)
		if err != nil {
			return err
		}
		tree.AddChild(n.Root())
	}

	// 更新 slide1.xml 文件
	err = utils.WriteXml(document, xmlFile)
	if err != nil {
		return err
	}

	// 4、压缩文件夹，生成新的 pptx 文件
	err = utils.CompressZip(tempDir, dstFile)
	if err != nil {
		return err
	}

	// 5、删除临时目录
	err = os.RemoveAll(tempDir)
	if err != nil {
		return err
	}
	return err
}

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

// GenTracerXLSX 生成可追踪表格
func GenTracerXLSX(srcFile, dstFile, traceUrl string) (err error) {
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
	tempDir, err = utils.ExtractZip(srcFile, "file-trace-*")
	if err != nil {
		return err
	}

	// 2、添加图片
	mediaDir = filepath.Join(tempDir, "xl", "media")
	err = utils.CreateDir(mediaDir)
	if err != nil {
		return err
	}

	count, err = utils.SubFileCount(mediaDir)
	if err != nil {
		return err
	}

	xlsxImageId = fmt.Sprintf("rId%d", count+1)
	xlsxImageName = fmt.Sprintf("image%d.png", count+1)
	err = os.WriteFile(filepath.Join(mediaDir, xlsxImageName), assets.MSMarkImage, os.ModePerm)
	if err != nil {
		return err
	}

	// 3、添加/修改 drawing1.xml，drawing1.xml.rels 文件
	drawingsDir = filepath.Join(tempDir, "xl", "drawings")
	drawingsRelsDir = filepath.Join(tempDir, "xl", "drawings", "_rels")
	err = utils.CreateDir(drawingsRelsDir)
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
		document, err = utils.ReadXml(typesFile)
		if err != nil {
			return err
		}

		types := document.SelectElement("Types")
		node := types.CreateElement("Default")
		node.CreateAttr("Extension", "png")
		node.CreateAttr("ContentType", "image/png")
		types.AddChild(node)

		// 更新 [Content_Types].xml 文件
		err = utils.WriteXml(document, typesFile)
		if err != nil {
			return err
		}
	} else {
		// 读取文件
		document, err = utils.ReadXml(relsFile)
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
		err = utils.WriteXml(document, relsFile)
		if err != nil {
			return err
		}
	}

	if _, err = os.Stat(xmlFile); err != nil {
		if count == 0 {
			// 创建文件
			// 添加追踪信息
			err = os.WriteFile(xmlFile, assets.MSDrawing, os.ModePerm)
			if err != nil {
				return err
			}

			// 添加 drawing1.xml 到 [Content_Types].xml 文件
			typesFile := filepath.Join(tempDir, "[Content_Types].xml")
			document, err = utils.ReadXml(typesFile)
			if err != nil {
				return err
			}

			types := document.SelectElement("Types")
			node := types.CreateElement("Override")
			node.CreateAttr("PartName", "/xl/drawings/drawing1.xml")
			node.CreateAttr("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
			types.AddChild(node)

			// 更新 [Content_Types].xml 文件
			err = utils.WriteXml(document, typesFile)
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
			assets.MSDrawingTpl = strings.Replace(assets.MSDrawingTpl, "${id}", strconv.Itoa(count+1), -1)
			assets.MSDrawingTpl = ">" + assets.MSDrawingTpl + "</xdr:wsDr>"

			current := strings.Replace(string(xmlContent), "></xdr:wsDr>", assets.MSDrawingTpl, -1)

			err = os.WriteFile(xmlFile, []byte(current), os.ModePerm)
			if err != nil {
				return err
			}
		}
	}

	// 4、添加/修改 sheet1.xml，sheet1.xml.rels 文件
	sheetDir = filepath.Join(tempDir, "xl", "worksheets")
	sheetRelsDir = filepath.Join(tempDir, "xl", "worksheets", "_rels")
	err = utils.CreateDir(sheetRelsDir)
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
		document, err = utils.ReadXml(xmlFile)
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
			err = utils.WriteXml(document, xmlFile)
			if err != nil {
				return err
			}
		}
	}

	// 5、压缩文件夹，生成新的 xlsx 文件
	err = utils.CompressZip(tempDir, dstFile)
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
