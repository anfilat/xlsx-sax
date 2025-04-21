package xlsx

import (
	"encoding/xml"
	"io"
)

func readStyleSheet(reader io.Reader) (*styleSheet, error) {
	decoder := xml.NewDecoder(reader)
	data := &styleSheet{}
	err := decoder.Decode(data)
	if err != nil {
		return nil, err
	}

	if data.NumFmts != nil {
		data.numFormats = make(map[int]string, len(data.NumFmts.NumFmt))

		for _, format := range data.NumFmts.NumFmt {
			data.numFormats[format.NumFmtId] = format.FormatCode
		}
	}

	return data, nil
}

type styleSheet struct {
	XMLName xml.Name `xml:"styleSheet"`
	NumFmts *struct {
		Count  int `xml:"count,attr"`
		NumFmt []struct {
			NumFmtId   int    `xml:"numFmtId,attr"`
			FormatCode string `xml:"formatCode,attr"`
		} `xml:"numFmt"`
	} `xml:"numFmts"`
	CellXfs struct {
		Count int `xml:"count,attr"`
		Xf    []struct {
			NumFmtId int `xml:"numFmtId,attr"`
		} `xml:"xf"`
	} `xml:"cellXfs"`

	numFormats map[int]string
}
