package main

import (
	"testing"
)

func TestExtractZip(t *testing.T) {
}

func TestCompressDir(t *testing.T) {
}

func TestReadXml(t *testing.T) {
}

func TestWriteXml(t *testing.T) {
}

func TestCreateDir(t *testing.T) {
	dir := "D:\\temp\\media\\test"
	err := CreateDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	t.Log("ok")
}

func TestSubFileCount(t *testing.T) {
	count, err := SubFileCount("assets")
	if err != nil {
		t.Fatal(err)
	}
	t.Log(count)
}
