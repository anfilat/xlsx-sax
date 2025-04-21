package xlsx

import "errors"

var (
	ErrWorkbookRelsNotExist = errors.New("parse xlsx file failed: xl/_rels/workbook.xml.rels doesn't exist")
	ErrWorkbookNotExist     = errors.New("parse xlsx file failed: xl/workbook.xml doesn't exist")
	ErrSheetNotFound        = errors.New("sheet RID doesn't found")
	ErrSheetNotExist        = errors.New("sheet doesn't exist")
)
