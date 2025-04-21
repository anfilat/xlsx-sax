package xlsx

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
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

	sheets, err := x.getWorkbookRels(workbookRelsFile)
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

	sharedStringFile := x.findFile(files, "sharedStrings.xml")
	if sharedStringFile != nil {
		err = x.fillSharedStrings(sharedStringFile)
		if err != nil {
			return err
		}
	}

	stylesFile := x.findFile(files, "styles.xml")
	if stylesFile != nil {
		err = x.fillStyles(stylesFile)
		if err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) findFile(files map[string]*zip.File, name string) *zip.File {
	for _, file := range files {
		if strings.HasSuffix(file.Name, name) {
			return file
		}
	}
	return nil
}

func (x *Xlsx) getWorkbookRels(zipFile *zip.File) (map[string]string, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	rels, err := readWorkbookRels(reader)
	if err != nil {
		return nil, err
	}

	sheets := make(map[string]string, len(rels.Relationship))
	for _, rel := range rels.Relationship {
		if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" {
			if strings.HasPrefix(rel.Target, "/xl/") {
				sheets[rel.ID] = rel.Target[1:]
			} else {
				sheets[rel.ID] = "xl/" + rel.Target
			}
		}
	}

	return sheets, nil
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
			continue
		}

		file, ok := files[path]
		if !ok {
			continue
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
