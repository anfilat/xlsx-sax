package xlsx

import (
	"encoding/xml"
	"io"
)

func readWorkbook(rd io.Reader) (*workbook, error) {
	decoder := xml.NewDecoder(rd)
	data := &workbook{}
	err := decoder.Decode(data)
	if err != nil {
		return nil, err
	}
	return data, nil
}

type workbook struct {
	XMLName xml.Name `xml:"workbook"`
	Sheets  []struct {
		Name    string `xml:"name,attr"`
		SheetId string `xml:"sheetId,attr"`
		ID      string `xml:"id,attr"`
	} `xml:"sheets>sheet"`
}
