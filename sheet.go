package xlsx

import (
	"archive/zip"
	"io"
	"strconv"
	"strings"

	"github.com/anfilat/xlsx-sax/xml"
)

type Sheet struct {
	zipReader     io.ReadCloser
	decoder       *xml.Decoder
	cols          []bool
	colIndex      []int
	countCols     int
	sharedStrings []string
}

func newSheetReader(zipFile *zip.File, cols []bool, skip int, sharedStrings []string) (*Sheet, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}

	countCols := 0
	colIndex := make([]int, len(cols))
	for i, value := range cols {
		if value {
			colIndex[i] = countCols
			countCols++
		}
	}

	decoder := xml.NewDecoder(reader)
	sheet := &Sheet{
		zipReader:     reader,
		decoder:       decoder,
		cols:          cols,
		colIndex:      colIndex,
		countCols:     countCols,
		sharedStrings: sharedStrings,
	}

	err = sheet.skipToSheetData()
	if err != nil {
		_ = reader.Close()
		return nil, err
	}

	for i := 0; i < skip; i++ {
		if !sheet.Next() {
			break
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

func (s *Sheet) Read() ([]string, error) {
	result := make([]string, s.countCols)

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
				return result, nil
			}
		case xml.CharData:
			if !isV {
				break
			}

			if cellName == "" {
				return nil, ErrIncorrectSheet
			}

			columnName := strings.TrimRight(cellName, "0123456789")
			ci := columnIndex(columnName)
			if ci >= len(s.cols) || !s.cols[ci] {
				break
			}

			val := string(token)
			if isSharedString {
				idx, err := strconv.Atoi(val)
				if err != nil {
					return nil, err
				}
				val = s.sharedStrings[idx]
			}

			result[s.colIndex[ci]] = val

			isV = false
		}
	}

	return nil, io.EOF
}
