package xlsx

import (
	"io"
	"strconv"

	"github.com/anfilat/xlsx-sax/internal/xml"
)

type styleSheet struct {
	numFormats    map[int]string
	cellXfs       []int
	parsedFormats map[string]*parsedNumFormat
}

func readStyleSheet(reader io.Reader) (*styleSheet, error) {
	decoder := xml.NewDecoder(reader, []xml.TagAttrs{
		{
			Name: "numFmt",
			Attr: []xml.TagAttr{
				{Name: "formatCode"},
				{Name: "numFmtId"},
			},
		},
		{
			Name: "xf",
			Attr: []xml.TagAttr{
				{Name: "numFmtId"},
			},
		},
	})

	result := styleSheet{
		numFormats:    make(map[int]string),
		parsedFormats: make(map[string]*parsedNumFormat),
	}

	isNumFmts := false
	isCellXfs := false
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch t.Type {
		case xml.StartElement:
			switch t.Name.Local {
			case "numFmts":
				isNumFmts = true
			case "numFmt":
				if isNumFmts {
					id := 0
					code := ""
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "formatCode":
							code = attr.Value.String()
						case "numFmtId":
							id, err = strconv.Atoi(string(attr.Value.Bytes()))
							if err != nil {
								return nil, err
							}
						}
					}

					if id > builtinNumFormatsCount {
						result.numFormats[id] = code
					}
				}
			case "cellXfs":
				isCellXfs = true
			case "xf":
				if isCellXfs {
					id := 0
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "numFmtId":
							id, err = strconv.Atoi(string(attr.Value.Bytes()))
							if err != nil {
								return nil, err
							}
						}
					}
					result.cellXfs = append(result.cellXfs, id)
				}
			case "styleSheet":
				//
			default:
				_ = decoder.Skip()
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "numFmts":
				isNumFmts = false
			case "cellXfs":
				isCellXfs = false
			}
		}
	}

	return &result, nil
}

func (s *styleSheet) getFormat(idx int) *parsedNumFormat {
	code := ""
	if idx >= 0 && idx < len(s.cellXfs) {
		xf := s.cellXfs[idx]
		if xf >= 0 && xf <= builtinNumFormatsCount {
			code = builtinNumFormats[xf]
		} else {
			code = s.numFormats[xf]
		}
	}

	if code == "" {
		code = "general"
	}

	format, ok := s.parsedFormats[code]
	if !ok {
		format = parseFullNumberFormatString(code)
		s.parsedFormats[code] = format
	}

	return format
}
