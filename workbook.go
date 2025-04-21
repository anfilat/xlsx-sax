package xlsx

import (
	"encoding/xml"
	"io"
)

func readWorkbook(reader io.Reader) (*workbook, error) {
	decoder := xml.NewDecoder(reader)
	data := &workbook{}
	err := decoder.Decode(data)
	if err != nil {
		return nil, err
	}
	return data, nil
}

type workbook struct {
	XMLName    xml.Name `xml:"workbook"`
	WorkbookPr struct {
		Date1904 bool `xml:"date1904,attr"`
	} `xml:"workbookPr"`
	Sheets []struct {
		Name    string `xml:"name,attr"`
		SheetId string `xml:"sheetId,attr"`
		ID      string `xml:"id,attr"`
	} `xml:"sheets>sheet"`
}
