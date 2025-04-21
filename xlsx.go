package xlsx

import (
	"archive/zip"
	"bytes"
)

type Xlsx struct {
	zip           *zip.Reader
	sharedStrings []string
}

type Params struct {
	Data []byte
}

func New(params Params) (*Xlsx, error) {
	br := bytes.NewReader(params.Data)
	zipReader, err := zip.NewReader(br, br.Size())
	if err != nil {
		return nil, err
	}

	result := Xlsx{
		zip: zipReader,
	}

	err = result.load()
	if err != nil {
		return nil, err
	}

	return &result, nil
}

func (x *Xlsx) load() error {
	for _, file := range x.zip.File {
		switch file.Name {
		case "xl/sharedStrings.xml":
			err := x.readSharedStrings(file)
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func (x *Xlsx) readSharedStrings(file *zip.File) error {
	reader, err := file.Open()
	if err != nil {
		return err
	}
	defer reader.Close()

	x.sharedStrings, err = readSharedStrings(reader)
	if err != nil {
		return err
	}
	return nil
}
