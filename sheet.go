package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strconv"
	"strings"
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

func (s *Sheet) Read(row *[]string) error {
	isV := false
	isSharedString := false
	cellName := ""
	for t, err := s.decoder.Token(); err == nil; t, err = s.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "c":
				isSharedString = false
				cellName = ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "t":
						isSharedString = a.Value == "s"
					case "r":
						cellName = a.Value
					}
				}
			case "v":
				isV = true
			}
		case xml.EndElement:
			isV = false
			if token.Name.Local == "row" {
				return nil
			}
		case xml.CharData:
			if !isV {
				break
			}

			if cellName == "" {
				return ErrIncorrectSheet
			}

			columnName := strings.TrimRight(cellName, "0123456789")
			_ = columnIndex(columnName)

			val := string(token)
			if isSharedString {
				idx, err := strconv.Atoi(val)
				if err != nil {
					return err
				}
				val = s.sharedStrings[idx]
			}

			*row = append(*row, val)

			isV = false
		}
	}

	return io.EOF
}
