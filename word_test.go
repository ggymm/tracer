package main

import (
	"github.com/beevik/etree"
	"os"
	"path/filepath"
	"testing"
)

func TestChangeXml(t *testing.T) {
	var (
		err     error
		tempDir string

		exist    = false
		document *etree.Document
	)

	// 解压文件
	tempDir, err = ExtractZip("testdata/file.docx", "file-trace-*")
	if err != nil {
		t.Fatal(err)
	}

	filename := filepath.Join(tempDir, "word", "settings.xml")
	document, err = ReadXml(filename)
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

func TestGenerateTracerDocx(t *testing.T) {
	srcFile := "example/source.docx"
	dstFile := "example/tracer.docx"
	traceUrl := "http://localhost:9090/trace"
	err := GenerateTracerDocx(srcFile, dstFile, traceUrl)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestGenerateTracerDocx2(t *testing.T) {
	srcFile := "example/tracer.docx"
	dstFile := "example/tracer2.docx"
	traceUrl := "http://localhost:9090/trace"
	err := GenerateTracerDocx(srcFile, dstFile, traceUrl)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}
