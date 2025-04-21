package xlsx

import (
	"encoding/xml"
	"io"
	"strconv"
)

func readSharedStrings(reader io.Reader) ([]string, error) {
	decoder := xml.NewDecoder(reader)

	var result []string
	isT := false
	isR := false
	index := 0
	for t, err := decoder.Token(); err == nil; t, err = decoder.Token() {
		switch token := t.(type) {
		case xml.StartElement:
			switch token.Name.Local {
			case "si":
			//
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
					result = make([]string, uniqCount)
				} else {
					result = make([]string, count)
				}
			default:
				_ = decoder.Skip()
			}
		case xml.EndElement:
			switch token.Name.Local {
			case "si":
				index++
			case "t":
				isT = false
			case "r":
				isR = false
			}
		case xml.CharData:
			if isT {
				if isR {
					result[index] += string(token)
				} else {
					result[index] = string(token)
				}
			}
		}
	}
	return result, nil
}
