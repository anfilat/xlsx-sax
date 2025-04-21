package xlsx

import (
	"archive/zip"
	"io"
	"strconv"

	"github.com/anfilat/xlsx-sax/internal/xml"
)

type Sheet struct {
	zipReader     io.ReadCloser
	decoder       *xml.Decoder
	sharedStrings sharedStrings
	styles        *styleSheet
	err           error

	cellValue      []byte
	isSharedString bool

	Row int
	Col int
}

func newSheetReader(zipFile *zip.File, sharedStrings sharedStrings, styles *styleSheet) (*Sheet, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}

	decoder := xml.NewDecoder(reader)
	sheet := &Sheet{
		zipReader:     reader,
		decoder:       decoder,
		sharedStrings: sharedStrings,
		styles:        styles,
	}

	err = sheet.skipToSheetData()
	if err != nil {
		_ = reader.Close()
		return nil, err
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
				if er := s.decoder.Skip(); er != nil {
					return er
				}
			}
		}
	}
	return nil
}

func (s *Sheet) Close() error {
	return s.zipReader.Close()
}

func (s *Sheet) Err() error {
	return s.err
}

func (s *Sheet) SkipRow() error {
	s.NextRow()
	return s.err
}

func (s *Sheet) NextRow() bool {
	if s.err != nil {
		return false
	}

	t, err := s.decoder.Token()
	for err == nil {
		switch token := t.(type) {
		case xml.StartElement:
			if token.Name.Local == "row" {
				for _, a := range token.Attr {
					if a.Name.Local == "r" {
						row, er := strconv.Atoi(a.Value)
						if er != nil {
							s.err = er
							return false
						}
						s.Row = row - 1
						s.Col = 0
						break
					}
				}
				return true
			}
		case xml.EndElement:
			if token.Name.Local == "sheetData" {
				return false
			}
		}

		t, err = s.decoder.Token()
	}

	s.err = err
	return false
}

func (s *Sheet) NextCell() bool {
	s.isSharedString = false
	isV := false

	t, err := s.decoder.Token()
	for err == nil {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "c":
				cellName := ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "t":
						s.isSharedString = a.Value == "s"
					case "r":
						cellName = a.Value
					}
				}
				if cellName == "" {
					s.err = ErrIncorrectSheet
					return false
				}

				s.Col = columnIndex(cellName)
			case "v":
				isV = true
			}
		case xml.CharData:
			if !isV {
				break
			}

			s.cellValue = token

			return true
		}

		t, err = s.decoder.Token()
	}

	s.err = err
	return false
}

func (s *Sheet) CellValue() (string, error) {
	if s.isSharedString {
		idx, err := strconv.Atoi(string(s.cellValue))
		if err != nil {
			return "", err
		}
		return s.sharedStrings.get(idx)
	}

	return string(s.cellValue), nil
}
