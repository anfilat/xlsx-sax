package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
)

type Sheet struct {
	zipReader     io.ReadCloser
	decoder       *xml.Decoder
	sharedStrings []string
}

type SheetParams struct {
	Skip int
}

func newSheetReader(zipFile *zip.File, params *SheetParams, sharedStrings []string) (*Sheet, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}

	decoder := xml.NewDecoder(reader)
	sheet := &Sheet{
		zipReader:     reader,
		decoder:       decoder,
		sharedStrings: sharedStrings,
	}

	err = sheet.skipToSheetData()
	if err != nil {
		_ = reader.Close()
		return nil, err
	}

	if params != nil && params.Skip > 0 {
		for i := 0; i < params.Skip; i++ {
			if !sheet.Next() {
				break
			}
		}
	}

	return sheet, nil
}

func (s *Sheet) skipToSheetData() error {
	for t, err := s.decoder.Token(); err == nil; t, err = s.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "worksheet":
				//
			case "sheetData":
				return nil
			default:
				if err := s.decoder.Skip(); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (s *Sheet) Close() error {
	return s.zipReader.Close()
}

func (s *Sheet) Next() bool {
	for t, err := s.decoder.Token(); err == nil; t, err = s.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "row":
				return true
			}
		}
	}
	return false
}
