package xlsx

import (
	"encoding/xml"
	"io"
)

func readStyles(reader io.Reader) (*styles, error) {
	decoder := xml.NewDecoder(reader)
	data := &styles{}
	err := decoder.Decode(data)
	if err != nil {
		return nil, err
	}
	return data, nil
}

type styles struct {
	XMLName xml.Name `xml:"styleSheet"`
}
