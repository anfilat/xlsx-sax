package xlsx

import (
	"archive/zip"
	"io"
)

type Xlsx struct {
	zip           *zip.Reader
	sheetNameFile map[string]*zip.File
	sheetIDFile   map[string]*zip.File
	sharedStrings []string
}

func New(reader io.ReaderAt, size int64) (*Xlsx, error) {
	zipReader, err := zip.NewReader(reader, size)
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

	sheets, sharedStringPath, err := x.getWorkbookRels(workbookRelsFile)
	if err != nil {
		return err
	}

	workbookFile, ok := files["xl/workbook.xml"]
	if !ok {
		return ErrWorkbookNotExist
	}

	err = x.fillWorkbook(workbookFile, sheets, files)
	if err != nil {
		return err
	}

	if sharedStringFile, ok := files[sharedStringPath]; ok {
		err = x.fillSharedStrings(sharedStringFile)
		if err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) getWorkbookRels(zipFile *zip.File) (map[string]string, string, error) {
	reader, err := zipFile.Open()
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

func (x *Xlsx) fillWorkbook(zipFile *zip.File, sheets map[string]string, files map[string]*zip.File) error {
	reader, err := zipFile.Open()
	if err != nil {
		return err
	}
	defer reader.Close()

	wb, err := readWorkbook(reader)
	if err != nil {
		return err
	}

	x.sheetNameFile = make(map[string]*zip.File, len(wb.Sheets.Sheet))
	x.sheetIDFile = make(map[string]*zip.File, len(wb.Sheets.Sheet))
	for _, sheet := range wb.Sheets.Sheet {
		path, ok := sheets[sheet.ID]
		if !ok {
			return ErrSheetNotFound
		}

		file, ok := files[path]
		if !ok {
			return ErrSheetNotExist
		}

		x.sheetNameFile[sheet.Name] = file
		x.sheetIDFile[sheet.SheetId] = file
	}

	return nil
}

func (x *Xlsx) fillSharedStrings(zipFile *zip.File) error {
	reader, err := zipFile.Open()
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
