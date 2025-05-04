package xlsx

import (
	"archive/zip"
	"io"
	"strconv"
	"time"

	"github.com/anfilat/xlsx-sax/internal/xml"
)

type Sheet struct {
	zipReader     io.ReadCloser
	decoder       *xml.Decoder
	sharedStrings sharedStrings
	styles        *styleSheet
	date1904      bool
	err           error

	isFutureRow bool
	futureRow   int

	cellValue  []byte
	cellType   cellType
	cellFormat int

	Row int
	Col int
}

type cellType int

const (
	cellTypeString cellType = iota
	cellTypeInline
	cellTypeFormula
	cellTypeBool
	cellTypeError
	cellTypeDate
	cellTypeNumeric
)

func newSheetReader(zipFile *zip.File, sharedStrings sharedStrings, styles *styleSheet, date1904 bool) (*Sheet, error) {
	reader, err := zipFile.Open()
	if err != nil {
		return nil, err
	}

	decoder := xml.NewDecoder(reader, []xml.TagAttrs{
		{
			Name: "row",
			Attr: []xml.TagAttr{
				{Name: "r"},
			},
		},
		{
			Name: "c",
			Attr: []xml.TagAttr{
				{Name: "t"},
				{Name: "s"},
				{Name: "r"},
			},
		},
	})
	sheet := &Sheet{
		zipReader:     reader,
		decoder:       decoder,
		sharedStrings: sharedStrings,
		styles:        styles,
		date1904:      date1904,
		cellValue:     make([]byte, 0),
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
		case *xml.StartElement:
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
	if s.NextRow() {
		for s.NextCell() {
		}
	}

	return s.err
}

func (s *Sheet) NextRow() bool {
	if s.err != nil {
		return false
	}

	if s.isFutureRow {
		s.isFutureRow = false
		s.Row = s.futureRow
		return true
	}

	row, err := s.nextRow()
	if err != nil {
		s.err = err
		return false
	}

	s.Row = row
	return true
}

func (s *Sheet) NextCell() bool {
	s.cellType = cellTypeNumeric
	s.cellFormat = 0

	isV := false
	isIs := false
	isT := false

	t, err := s.decoder.Token()
	for err == nil {
		switch token := t.(type) {
		case *xml.StartElement:
			switch token.Name.Local {
			case "c":
				var cell []byte
				for _, a := range token.Attr {
					switch a.Name.Local {
					case "t":
						switch string(a.Value.Bytes()) {
						case "s":
							s.cellType = cellTypeString
						case "inlineStr":
							s.cellType = cellTypeInline
						case "b":
							s.cellType = cellTypeBool
						case "e":
							s.cellType = cellTypeError
						case "str":
							s.cellType = cellTypeFormula
						case "d":
							s.cellType = cellTypeDate
						case "n":
							s.cellType = cellTypeNumeric
						}
					case "s":
						s.cellFormat, err = strconv.Atoi(a.Value.String())
						if err != nil {
							s.err = ErrIncorrectSheet
							return false
						}
					case "r":
						cell = a.Value.Bytes()
					}
				}
				if len(cell) == 0 {
					s.err = ErrIncorrectSheet
					return false
				}

				s.Col = columnIndex(cell)
				s.cellValue = s.cellValue[:0]
			case "v":
				isV = true
			case "is":
				isIs = true
			case "t":
				isT = true
			}
		case *xml.EndElement:
			switch token.Name.Local {
			case "c":
				return true
			case "row":
				row, er := s.nextRow()
				if er != nil {
					s.err = er
					return false
				}

				if row != s.Row {
					s.isFutureRow = true
					s.futureRow = row
					return false
				}
			case "v":
				isV = false
			case "is":
				isIs = false
			case "t":
				isT = false
			}
		case *xml.CharData:
			if !(isV || (isIs && isT)) {
				break
			}

			s.cellValue = append(s.cellValue, token.Value...)
		}

		t, err = s.decoder.Token()
	}

	s.err = err
	return false
}

func (s *Sheet) nextRow() (int, error) {
	t, err := s.decoder.Token()
	for err == nil {
		switch token := t.(type) {
		case *xml.StartElement:
			if token.Name.Local == "row" {
				row, er := parseRowNumber(token.Attr)
				if er != nil {
					return 0, er
				}
				return row - 1, nil
			}
		case *xml.EndElement:
			if token.Name.Local == "sheetData" {
				return 0, io.EOF
			}
		}

		t, err = s.decoder.Token()
	}

	return 0, err
}

func parseRowNumber(attrs []xml.Attr) (int, error) {
	for _, attr := range attrs {
		if attr.Name.Local == "r" {
			return strconv.Atoi(attr.Value.String())
		}
	}
	return 0, ErrRowMissingR
}

func (s *Sheet) CellValue() (string, error) {
	if s.cellType == cellTypeString {
		return s.getSharedString()
	}

	return string(s.cellValue), nil
}

func (s *Sheet) CellFloat() (float64, error) {
	if s.cellType == cellTypeString {
		str, err := s.getSharedString()
		if err != nil {
			return 0, err
		}

		return strconv.ParseFloat(str, 64)
	}

	return strconv.ParseFloat(string(s.cellValue), 64)
}

func (s *Sheet) CellInt() (int, error) {
	if s.cellType == cellTypeString {
		str, err := s.getSharedString()
		if err != nil {
			return 0, err
		}

		return strconv.Atoi(str)
	}

	return strconv.Atoi(string(s.cellValue))
}

func (s *Sheet) CellTime() (time.Time, error) {
	val, err := s.CellFloat()
	if err != nil {
		return time.Time{}, err
	}
	return timeFromExcelTime(val, s.date1904), nil
}

func (s *Sheet) CellFormatValue() (string, error) {
	switch s.cellType {
	case cellTypeString:
		format := s.styles.getFormat(s.cellFormat)
		str, err := s.getSharedString()
		if err != nil {
			return "", err
		}
		val, err := format.text(str)
		if format.parseEncounteredError != nil {
			return val, format.parseEncounteredError
		}
		return val, err
	case cellTypeInline, cellTypeFormula:
		format := s.styles.getFormat(s.cellFormat)
		val, err := format.text(string(s.cellValue))
		if format.parseEncounteredError != nil {
			return val, format.parseEncounteredError
		}
		return val, err
	case cellTypeBool:
		if string(s.cellValue) == "0" {
			return "FALSE", nil
		}
		if string(s.cellValue) == "1" {
			return "TRUE", nil
		}
		return string(s.cellValue), ErrInvalidBool
	case cellTypeError, cellTypeDate:
		return string(s.cellValue), nil
	case cellTypeNumeric:
		format := s.styles.getFormat(s.cellFormat)
		val, err := format.numeric(string(s.cellValue), s.date1904)
		if format.parseEncounteredError != nil {
			return val, format.parseEncounteredError
		}
		return val, err
	default:
		return string(s.cellValue), ErrUnknownCellType
	}
}

func (s *Sheet) getSharedString() (string, error) {
	idx, err := strconv.Atoi(string(s.cellValue))
	if err != nil {
		return "", err
	}

	return s.sharedStrings.get(idx)
}
