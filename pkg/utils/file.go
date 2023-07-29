package utils

import "os"

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
