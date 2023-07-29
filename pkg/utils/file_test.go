package utils

import "testing"

func TestCreateDir(t *testing.T) {
	type args struct {
		dir string
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
	}{
		{"test1", args{"test1"}, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err := CreateDir(tt.args.dir); (err != nil) != tt.wantErr {
				t.Errorf("CreateDir() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

func TestSubFileCount(t *testing.T) {
	type args struct {
		dir string
	}
	tests := []struct {
		name    string
		args    args
		wantNum int
		wantErr bool
	}{
		{"test1", args{"test1"}, 0, false},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotNum, err := SubFileCount(tt.args.dir)
			if (err != nil) != tt.wantErr {
				t.Errorf("SubFileCount() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if gotNum != tt.wantNum {
				t.Errorf("SubFileCount() gotNum = %v, want %v", gotNum, tt.wantNum)
			}
		})
	}
}
