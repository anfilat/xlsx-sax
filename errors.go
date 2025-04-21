package xlsx

import "errors"

var (
	ErrWorkbookRelsNotExist  = errors.New("parse xlsx file failed: xl/_rels/workbook.xml.rels doesn't exist")
	ErrWorkbookNotExist      = errors.New("parse xlsx file failed: xl/workbook.xml doesn't exist")
	ErrParseWorkbook         = errors.New("parse workbook")
	ErrSheetNotFound         = errors.New("sheet not found")
	ErrIncorrectSheet        = errors.New("incorrect sheet")
	ErrIncorrectSharedString = errors.New("incorrect shared string")
	ErrParseStyles           = errors.New("parse styles")
	ErrDoubleQuote           = errors.New("invalid format string, unmatched double quote")
	ErrManySections          = errors.New("invalid number format, too many format sections")
	ErrUnknownCellType       = errors.New("unknown cell type")
	ErrInvalidBool           = errors.New("invalid value in bool cell")
	ErrInvalidFormat         = errors.New("invalid or unsupported format")
)
