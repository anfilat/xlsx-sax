package xlsx

import (
	"archive/zip"
	"io"
	"sort"
	"strconv"

	"github.com/anfilat/xlsx-sax/internal/xml"
)

type Sheet struct {
	zipReader     io.ReadCloser
	decoder       *xml.Decoder
	colGet        []bool
	colMap        []int
	countCols     int
	sharedStrings []string
}

func newSheetReader(zipFile *zip.File, cols []int, skip int, sharedStrings []string) (*Sheet, error) {
	if len(cols) == 0 {
		return nil, ErrNoColumns
	}

	sort.Ints(cols)

	colGet := make([]bool, cols[len(cols)-1]+1)
	colMap := make([]int, cols[len(cols)-1]+1)
	countCols := 0
	for _, value := range cols {
		colGet[value] = true
		colMap[value] = countCols
		countCols++
	}

	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}

	decoder := xml.NewDecoder(reader)
	sheet := &Sheet{
		zipReader:     reader,
		decoder:       decoder,
		colGet:        colGet,
		colMap:        colMap,
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
			if token.Name.Local == "row" {
				return true
			}
		case xml.EndElement:
			if token.Name.Local == "sheetData" {
				return false
			}
		}
	}
	return false
}

func (s *Sheet) Read(row []string) error {
	isV := false
	isSharedString := false
	ci := -1
	for t, err := s.decoder.Token(); err == nil; t, err = s.decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "c":
				isSharedString = false
				ci = -1
				cellName := ""
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "t":
						isSharedString = a.Value == "s"
					case "r":
						cellName = a.Value
					}
				}
				if cellName == "" {
					return ErrIncorrectSheet
				}

				ci = columnIndex(cellName)
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

			if ci >= len(s.colGet) || !s.colGet[ci] {
				break
			}

			if isSharedString {
				idx, err := strconv.Atoi(string(token))
				if err != nil {
					return err
				}
				row[s.colMap[ci]] = s.sharedStrings[idx]
			} else {
				row[s.colMap[ci]] = string(token)
			}

			isV = false
		}
	}

	return io.EOF
}
