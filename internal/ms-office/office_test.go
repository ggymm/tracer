package ms_office

import (
	"github.com/beevik/etree"
	"os"
	"path/filepath"
	"testing"
	"tracer/internal/assets"
	"tracer/pkg/utils"
)

func TestChangeXml(t *testing.T) {
	var (
		err     error
		tempDir string

		exist    = false
		document *etree.Document
	)

	// 解压文件
	tempDir, err = utils.ExtractZip("C:/Product/deceptive-defense/tracer/example/file.docx", "file-trace-*")
	if err != nil {
		t.Fatal(err)
	}

	filename := filepath.Join(tempDir, "word", "settings.xml")
	document, err = utils.ReadXml(filename)
	if err != nil {
		t.Fatal(err)
	}

	document.Indent(4)
	t.Log(document.WriteToString())

	root := document.SelectElement("w:settings")
	for _, element := range root.ChildElements() {
		if element.Space == "w" && element.Tag == "attachedTemplate" {
			exist = true
			break
		}
	}

	if !exist {
		// 添加节点
		attachedTemplate := root.CreateElement("w:attachedTemplate")
		attachedTemplate.CreateAttr("r:id", "rId1")

		root.InsertChildAt(1, attachedTemplate)
	}

	document.Indent(4)
	t.Log(document.WriteToString())

	// 移除临时目录
	err = os.RemoveAll(tempDir)
	if err != nil {
		t.Fatal(err)
	}
}

func TestGenTracerDOCX(t *testing.T) {
	srcFile := "C:/Product/deceptive-defense/tracer/example/source.docx"
	dstFile := "C:/Product/deceptive-defense/tracer/example/tracer.docx"
	traceUrl := "http://localhost:9090/trace"
	err := GenTracerDOCX(srcFile, dstFile, traceUrl)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestGenTracerDOCX2(t *testing.T) {
	srcFile := "C:/Product/deceptive-defense/tracer/example/tracer.docx"
	dstFile := "C:/Product/deceptive-defense/tracer/example/tracer2.docx"
	traceUrl := "http://localhost:9090/trace"
	err := GenTracerDOCX(srcFile, dstFile, traceUrl)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestWriteImage(t *testing.T) {
	err := os.WriteFile("D:\\temp\\image.png", assets.MSMarkImage, os.ModePerm)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestGenTracerPPTX(t *testing.T) {
	srcFile := "C:/Product/deceptive-defense/tracer/example/source.pptx"
	dstFile := "C:/Product/deceptive-defense/tracer/example/tracer.pptx"
	traceUrl := "http://localhost:9090/trace"
	t.Log(GenTracerPPTX(srcFile, dstFile, traceUrl))
}

func TestGenTracerXLSX(t *testing.T) {
	srcFile := "C:/Product/deceptive-defense/tracer/example/source.xlsx"
	dstFile := "C:/Product/deceptive-defense/tracer/example/tracer.xlsx"
	traceUrl := "http://localhost:9090/trace"
	t.Log(GenTracerXLSX(srcFile, dstFile, traceUrl))
}

func TestGenTracerXLSX2(t *testing.T) {
	srcFile := "C:/Product/deceptive-defense/tracer/example/source-image.xlsx"
	dstFile := "C:/Product/deceptive-defense/tracer/example/tracer-image.xlsx"
	traceUrl := "http://localhost:9090/trace"
	t.Log(GenTracerXLSX(srcFile, dstFile, traceUrl))
}
