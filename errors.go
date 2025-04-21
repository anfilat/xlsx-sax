package xlsx

import "errors"

var (
	ErrWorkbookRelsNotExist  = errors.New("parse xlsx file failed: xl/_rels/workbook.xml.rels doesn't exist")
	ErrWorkbookNotExist      = errors.New("parse xlsx file failed: xl/workbook.xml doesn't exist")
	ErrParseWorkbook         = errors.New("parse workbook")
	ErrSheetNotFound         = errors.New("sheet not found")
	ErrIncorrectSheet        = errors.New("incorrect sheet")
	ErrIncorrectSharedString = errors.New("incorrect shared string")
	ErrDoubleQuote           = errors.New("invalid format string, unmatched double quote")
	ErrManySections          = errors.New("invalid number format, too many format sections")
	ErrInvalidBrackets       = errors.New("invalid formatting code, invalid brackets")
	ErrInvalidCurrency       = errors.New("invalid formatting code, invalid currency annotation")
	EUnsupportedCharacters   = errors.New("invalid formatting code: unsupported or unescaped characters")
	ErrUnknownCellType       = errors.New("unknown cell type")
	ErrInvalidBool           = errors.New("invalid value in bool cell")
	ErrInvalidFormat         = errors.New("invalid or unsupported format")
	ErrNoClosingQuote        = errors.New("no closing quote found")
)
