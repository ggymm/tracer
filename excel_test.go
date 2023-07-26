package main

import (
	"os"
	"testing"
)

func TestWriteImage(t *testing.T) {
	err := os.WriteFile("D:\\temp\\image.png", markImage, os.ModePerm)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestGenerateTracerXlsx(t *testing.T) {
	srcFile := "example/source.xlsx"
	dstFile := "example/tracer.xlsx"
	traceUrl := "http://localhost:9090/trace"
	t.Log(GenerateTracerXlsx(srcFile, dstFile, traceUrl))
}

func TestGenerateTracerXlsx2(t *testing.T) {
	srcFile := "example/source-image.xlsx"
	dstFile := "example/tracer-image.xlsx"
	traceUrl := "http://localhost:9090/trace"
	t.Log(GenerateTracerXlsx(srcFile, dstFile, traceUrl))
}
