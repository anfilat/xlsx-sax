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
	files := make(map[string]*zip.File, len(x.zip.File))
	for _, file := range x.zip.File {
		files[file.Name] = file
	}

	workbookRelsFile, ok := files["xl/_rels/workbook.xml.rels"]
	if !ok {
		return ErrWorkbookRelsNotExist
	}

	_, sharedStringPath, err := x.readWorkbookRels(workbookRelsFile)
	if err != nil {
		return err
	}

	sharedStringFile, ok := files[sharedStringPath]
	if ok {
		err = x.readSharedStrings(sharedStringFile)
		if err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) readWorkbookRels(file *zip.File) (map[string]string, string, error) {
	reader, err := file.Open()
	if err != nil {
		return nil, "", err
	}
	defer reader.Close()

	rels, err := readWorkbookRels(reader)
	if err != nil {
		return nil, "", err
	}

	sheets := make(map[string]string, len(rels.Relationship))
	var sharedStrings string
	for _, rel := range rels.Relationship {
		switch rel.Type {
		case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
			sheets[rel.ID] = "xl/" + rel.Target
		case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
			sharedStrings = "xl/" + rel.Target
		}
	}

	return sheets, sharedStrings, nil
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
