package xlsx

import (
	"io"
	"strconv"

	"github.com/anfilat/xlsx-sax/internal/xml"
)

type sharedStrings []string

func (s sharedStrings) get(idx int) (string, error) {
	if idx < 0 || idx >= len(s) {
		return "", ErrIncorrectSharedString
	}
	return s[idx], nil
}

func readSharedStrings(reader io.Reader) (sharedStrings, error) {
	decoder := xml.NewDecoder(reader)

	var result sharedStrings
	ar := &arena{}
	isT := false
	isR := false
	str := ""
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch token := t.(type) {
		case *xml.StartElement:
			switch token.Name.Local {
			case "si":
				str = ""
			case "t":
				isT = true
			case "r":
				isR = true
			case "sst":
				uniqCount := 0
				count := 0
				for _, attr := range token.Attr {
					switch attr.Name.Local {
					case "uniqueCount":
						uniqCount, err = strconv.Atoi(attr.Value)
						if err != nil {
							return nil, err
						}
					case "count":
						count, err = strconv.Atoi(attr.Value)
						if err != nil {
							return nil, err
						}
					}
				}
				if uniqCount != 0 {
					result = make([]string, 0, uniqCount)
				} else {
					result = make([]string, 0, count)
				}
			default:
				_ = decoder.Skip()
			}
		case *xml.EndElement:
			switch token.Name.Local {
			case "si":
				result = append(result, str)
			case "t":
				isT = false
			case "r":
				isR = false
			}
		case *xml.CharData:
			if isT {
				if isR {
					str += ar.toString(token.Value)
				} else {
					str = ar.toString(token.Value)
				}
			}
		}
	}
	return result, nil
}
