package xlsx

import (
	"encoding/xml"
	"io"
)

func readWorkbookRels(reader io.Reader) (*workbookRels, error) {
	decoder := xml.NewDecoder(reader)
	data := &workbookRels{}
	err := decoder.Decode(data)
	if err != nil {
		return nil, err
	}
	return data, nil
}

type workbookRels struct {
	XMLName      xml.Name `xml:"Relationships"`
	Relationship []struct {
		ID     string `xml:"Id,attr"`
		Type   string `xml:"Type,attr"`
		Target string `xml:"Target,attr"`
	} `xml:"Relationship"`
}
