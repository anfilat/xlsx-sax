package xlsx

import (
	"archive/zip"
	"fmt"
	"io"
)

type Xlsx struct {
	zip           *zip.Reader
	date1904      bool
	sheetFile     []*zip.File
	sheetNames    []string
	sheetNameFile map[string]*zip.File
	sharedStrings sharedStrings
	styles        *styleSheet
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

	sheets, sharedStringPath, stylesPath, err := x.getWorkbookRels(workbookRelsFile)
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

	sharedStringFile, ok := files[sharedStringPath]
	if ok {
		err = x.fillSharedStrings(sharedStringFile)
		if err != nil {
			return err
		}
	}

	stylesFile, ok := files[stylesPath]
	if ok {
		err = x.fillStyles(stylesFile)
		if err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) getWorkbookRels(zipFile *zip.File) (map[string]string, string, string, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, "", "", err
	}
	defer reader.Close()

	rels, err := readWorkbookRels(reader)
	if err != nil {
		return nil, "", "", err
	}

	sheets := make(map[string]string, len(rels.Relationship))
	var sharedStrs string
	var styles string
	for _, rel := range rels.Relationship {
		switch rel.Type {
		case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
			sheets[rel.ID] = "xl/" + rel.Target
		case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
			sharedStrs = "xl/" + rel.Target
		case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
			styles = "xl/" + rel.Target
		}
	}

	return sheets, sharedStrs, styles, nil
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

	x.date1904 = wb.WorkbookPr.Date1904

	x.sheetFile = make([]*zip.File, 0, len(wb.Sheets))
	x.sheetNames = make([]string, 0, len(wb.Sheets))
	x.sheetNameFile = make(map[string]*zip.File, len(wb.Sheets))
	for _, sheet := range wb.Sheets {
		path, ok := sheets[sheet.ID]
		if !ok {
			return fmt.Errorf("sheet RID %s doesn't found: %w", sheet.ID, ErrParseWorkbook)
		}

		file, ok := files[path]
		if !ok {
			return fmt.Errorf("sheet %s doesn't exist: %w", path, ErrParseWorkbook)
		}

		x.sheetFile = append(x.sheetFile, file)
		x.sheetNames = append(x.sheetNames, sheet.Name)
		x.sheetNameFile[sheet.Name] = file
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

func (x *Xlsx) fillStyles(zipFile *zip.File) error {
	reader, err := zipFile.Open()
	if err != nil {
		return err
	}
	defer reader.Close()

	x.styles, err = readStyleSheet(reader)
	if err != nil {
		return err
	}
	return nil
}

func (x *Xlsx) SheetNames() []string {
	result := make([]string, len(x.sheetNames))
	copy(result, x.sheetNames)
	return result
}

func (x *Xlsx) OpenSheetByName(name string) (*Sheet, error) {
	file, ok := x.sheetNameFile[name]
	if !ok {
		return nil, fmt.Errorf("can not find worksheet %s: %w", name, ErrSheetNotFound)
	}

	return newSheetReader(file, x.sharedStrings, x.styles, x.date1904)
}

func (x *Xlsx) OpenSheetByOrder(n int) (*Sheet, error) {
	if n < 0 || n >= len(x.sheetFile) {
		return nil, fmt.Errorf("can not find worksheet %d: %w", n, ErrSheetNotFound)
	}

	file := x.sheetFile[n]
	return newSheetReader(file, x.sharedStrings, x.styles, x.date1904)
}
