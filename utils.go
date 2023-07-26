package main

import (
	"archive/zip"
	"github.com/beevik/etree"
	"io"
	"os"
	"path/filepath"
)

// ExtractZip 解压 zip 文件
// filename: 压缩文件名
// patten: 临时目录名
func ExtractZip(filename string, patten string) (tempDir string, err error) {
	var (
		zipFile *zip.ReadCloser
	)

	if patten == "" {
		patten = "temp-zip-*"
	}

	// 创建临时目录
	tempDir, err = os.MkdirTemp("", patten)
	if err != nil {
		return tempDir, err
	}

	// 读取压缩文件
	zipFile, err = zip.OpenReader(filename)
	if err != nil {
		return tempDir, err
	}

	for _, file := range zipFile.Reader.File {
		// 忽略目录
		if file.FileInfo().IsDir() {
			continue
		}

		var (
			dstPath = filepath.Join(tempDir, file.Name)

			dstFile *os.File
			srcFile io.ReadCloser
		)

		// 如果文件目录不存在，则创建
		dstDir := filepath.Dir(dstPath)
		if _, err = os.Stat(dstDir); err != nil {
			if err = os.MkdirAll(dstDir, os.ModePerm); err != nil {
				return tempDir, err
			}
		}

		// 创建目标文件
		dstFile, err = os.OpenFile(dstPath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, file.Mode())
		if err != nil {
			return tempDir, err
		}

		// 打开源文件
		srcFile, err = file.Open()
		if err != nil {
			return tempDir, err
		}

		// 拷贝文件
		_, err = io.Copy(dstFile, srcFile)
		if err != nil {
			return tempDir, err
		}

		// 关闭文件
		_ = srcFile.Close()
		_ = dstFile.Close()
	}

	return tempDir, nil
}

// CompressZip 压缩文件夹
// dir: 目录名
// filename: 压缩文件名
func CompressZip(dir string, filename string) (err error) {
	var (
		dstFile   *os.File
		dstWriter *zip.Writer
	)

	// 创建压缩文件
	dstFile, err = os.Create(filename)
	if err != nil {
		return err
	}
	defer func() {
		_ = dstFile.Close()
	}()

	// 创建压缩器
	dstWriter = zip.NewWriter(dstFile)

	// 遍历目录
	err = filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		var (
			dst io.Writer

			relPath string
			relFile *os.File
		)

		// 获取相对路径
		relPath, err = filepath.Rel(dir, path)
		if err != nil {
			return err
		}

		if relPath == "." {
			return nil
		}

		if info.IsDir() {
			// 创建目录
			_, err = dstWriter.Create(relPath + "/")
			if err != nil {
				return err
			}
		} else {
			// 创建文件
			relFile, err = os.Open(path)
			if err != nil {
				return err
			}

			dst, err = dstWriter.Create(relPath)
			if err != nil {
				return err
			}

			// 拷贝文件
			_, err = io.Copy(dst, relFile)
			if err != nil {
				return err
			}

			_ = relFile.Close()
		}

		return nil
	})
	if err != nil {
		return err
	}

	// 关闭压缩器
	// 必须最后关闭
	err = dstWriter.Close()
	if err != nil {
		return err
	}

	return nil
}

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

func CreateDir(dir string) (err error) {
	if _, err = os.Stat(dir); err != nil {
		err = os.MkdirAll(dir, os.ModePerm)
		if err != nil {
			return err
		}
	}

	return nil
}

func SubFileCount(dir string) (num int, err error) {
	var fileInfoList []os.DirEntry
	fileInfoList, err = os.ReadDir(dir)
	if err != nil {
		return 0, err
	}

	return len(fileInfoList), nil
}
