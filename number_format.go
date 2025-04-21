package xlsx

// Most of this file was taken from https://github.com/tealeg/xlsx

import (
	"errors"
	"fmt"
	"strings"
)

const builtinNumFormatsCount = 163

var builtinNumFormats = []string{
	0x00: `General`,
	0x01: `0`,
	0x02: `0.00`,
	0x03: `#,##0`,
	0x04: `#,##0.00`,
	0x05: `($#,##0_);($#,##0)`,
	0x06: `($#,##0_);[RED]($#,##0)`,
	0x07: `($#,##0.00_);($#,##0.00_)`,
	0x08: `($#,##0.00_);[RED]($#,##0.00_)`,
	0x09: `0%`,
	0x0a: `0.00%`,
	0x0b: `0.00E+00`,
	0x0c: `# ?/?`,
	0x0d: `# ??/??`,
	0x0e: `m-d-yy`,
	0x0f: `d-mmm-yy`,
	0x10: `d-mmm`,
	0x11: `mmm-yy`,
	0x12: `h:mm AM/PM`,
	0x13: `h:mm:ss AM/PM`,
	0x14: `h:mm`,
	0x15: `h:mm:ss`,
	0x16: `m-d-yy h:mm`,
	0x25: `(#,##0_);(#,##0)`,
	0x26: `(#,##0_);[RED](#,##0)`,
	0x27: `(#,##0.00);(#,##0.00)`,
	0x28: `(#,##0.00);[RED](#,##0.00)`,
	0x29: `_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)`,
	0x2a: `_($*#,##0_);_($*(#,##0);_(*"-"_);_(@_)`,
	0x2b: `_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)`,
	0x2c: `_($*#,##0.00_);_($*(#,##0.00);_(*"-"??_);_(@_)`,
	0x2d: `mm:ss`,
	0x2e: `[h]:mm:ss`,
	0x2f: `mm:ss.0`,
	0x30: `##0.0E+0`,
	0x31: `@`,
}

type parsedNumFormat struct {
	numFmt                        string
	positiveFormat                *formatOptions
	negativeFormat                *formatOptions
	zeroFormat                    *formatOptions
	textFormat                    *formatOptions
	parseEncounteredError         error
	isTimeFormat                  bool
	negativeFormatExpectsPositive bool
}

type formatOptions struct {
	fullFormatString    string
	reducedFormatString string
	prefix              string
	suffix              string
	isTimeFormat        bool
	showPercent         bool
}

func parseFullNumberFormatString(numFmt string) *parsedNumFormat {
	parsedNumFmt := &parsedNumFormat{
		numFmt: numFmt,
	}

	var fmtOptions []*formatOptions
	formats, err := splitFormat(numFmt)
	if err == nil {
		for _, formatSection := range formats {
			parsedFormat, err := parseNumberFormatSection(formatSection)
			if err != nil {
				parsedFormat = fallbackErrorFormat
				parsedNumFmt.parseEncounteredError = err
			}
			fmtOptions = append(fmtOptions, parsedFormat)
		}
	} else {
		fmtOptions = append(fmtOptions, fallbackErrorFormat)
		parsedNumFmt.parseEncounteredError = err
	}

	if len(fmtOptions) > 4 {
		fmtOptions = []*formatOptions{fallbackErrorFormat}
		parsedNumFmt.parseEncounteredError = ErrManySections
	}

	if len(fmtOptions) == 1 {
		// If there is only one option, it is used for all
		parsedNumFmt.isTimeFormat = fmtOptions[0].isTimeFormat
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[0]
		parsedNumFmt.zeroFormat = fmtOptions[0]
		if strings.Contains(fmtOptions[0].fullFormatString, "@") {
			parsedNumFmt.textFormat = fmtOptions[0]
		} else {
			parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
		}
	} else if len(fmtOptions) == 2 {
		// If there are two formats, the first is used for positive and zeros, the second gets used as a negative format,
		// and strings are not formatted.
		// When negative numbers now have their own format, they should become positive before having the format applied.
		// The format will contain a negative sign if it is desired, but they may be colored red or wrapped in
		// parenthesis instead.
		parsedNumFmt.isTimeFormat = fmtOptions[0].isTimeFormat
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[0]
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
	} else if len(fmtOptions) == 3 {
		// If there are three formats, the first is used for positive, the second gets used as a negative format,
		// the third is for negative, and strings are not formatted.
		parsedNumFmt.isTimeFormat = fmtOptions[0].isTimeFormat
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[2]
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
	} else {
		// With four options, the first is positive, the second is negative, the third is zero, and the fourth is strings
		// Negative numbers should be still become positive before having the negative formatting applied.
		parsedNumFmt.isTimeFormat = fmtOptions[0].isTimeFormat
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[2]
		parsedNumFmt.textFormat = fmtOptions[3]
	}
	return parsedNumFmt
}

var fallbackErrorFormat = &formatOptions{
	fullFormatString:    "general",
	reducedFormatString: "general",
}

func splitFormat(format string) ([]string, error) {
	var result []string
	prevIndex := 0
	for i := 0; i < len(format); i++ {
		if format[i] == ';' {
			result = append(result, format[prevIndex:i])
			prevIndex = i + 1
		} else if format[i] == '\\' {
			i++
		} else if format[i] == '"' {
			endQuoteIndex := strings.Index(format[i+1:], `"`)
			if endQuoteIndex == -1 {
				return nil, ErrDoubleQuote
			}
			i += endQuoteIndex + 1
		}
	}
	return append(result, format[prevIndex:]), nil
}

func parseNumberFormatSection(fullFormat string) (*formatOptions, error) {
	reducedFormat := strings.TrimSpace(fullFormat)

	// general is the only format that does not use the normal format symbols notations
	if compareFormatString(reducedFormat, "general") {
		return &formatOptions{
			fullFormatString:    "general",
			reducedFormatString: "general",
		}, nil
	}
	if isTimeFormat(reducedFormat) {
		return &formatOptions{
			fullFormatString:    fullFormat,
			isTimeFormat:        true,
			reducedFormatString: reducedFormat,
		}, nil
	}

	prefix, reducedFormat, showPercent1, err := parseLiterals(reducedFormat)
	if err != nil {
		return nil, err
	}

	reducedFormat, suffixFormat := splitFormatAndSuffixFormat(reducedFormat)

	suffix, remaining, showPercent2, err := parseLiterals(suffixFormat)
	if err != nil {
		return nil, err
	}
	if len(remaining) > 0 {
		// This paradigm of codes consisting of literals, number formats, then more literals is not always correct, they can
		// actually be intertwined. Though 99% of the time number formats will not do this.
		// Excel uses this format string for Social Security Numbers: 000\-00\-0000
		// and this for US phone numbers: [<=9999999]###\-####;\(###\)\ ###\-####
		return nil, errors.New("invalid or unsupported format string")
	}

	return &formatOptions{
		fullFormatString:    fullFormat,
		isTimeFormat:        false,
		reducedFormatString: reducedFormat,
		prefix:              prefix,
		suffix:              suffix,
		showPercent:         showPercent1 || showPercent2,
	}, nil
}

func compareFormatString(fmt1, fmt2 string) bool {
	if fmt1 == fmt2 {
		return true
	}
	if fmt1 == "" || strings.EqualFold(fmt1, "general") {
		fmt1 = "general"
	}
	if fmt2 == "" || strings.EqualFold(fmt2, "general") {
		fmt2 = "general"
	}
	return fmt1 == fmt2
}

func isTimeFormat(format string) bool {
	var foundTimeFormatCharacters bool

	runes := []rune(format)
	for i := 0; i < len(runes); i++ {
		curReducedFormat := runes[i:]
		switch curReducedFormat[0] {
		case '\\', '_':
			// If there is a slash, skip the next character, and add it to the prefix
			// If there is an underscore, skip the next character, but don't add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
			// Asterisks are used to repeat the next character to fill the full cell width.
			// There isn't really a cell size in this context, so this will be ignored.
		case '"':
			// If there is a quote skip to the next quote, and add the quoted characters to the prefix
			endQuoteIndex, err := skipToRune(curReducedFormat, '"')
			if err != nil {
				return false
			}
			i += endQuoteIndex + 1
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ':
			// These symbols are allowed to be used as literal without escaping
		case ',':
			// This is not documented in the XLSX spec as far as I can tell, but Excel and Numbers will include
			// commas in number formats without escaping them, so this should be supported.
		default:
			foundInThisLoop := false
			for _, special := range timeFormatCharacters {
				if strings.HasPrefix(string(curReducedFormat), special) {
					foundTimeFormatCharacters = true
					foundInThisLoop = true
					i += len([]rune(special)) - 1
					break
				}
			}
			if foundInThisLoop {
				continue
			}
			if curReducedFormat[0] == '[' {
				// For number formats, this code would happen above in a case '[': section.
				// However, for time formats it must happen after looking for occurrences in timeFormatCharacters
				// because there are a few time formats that can be wrapped in brackets.

				// Brackets can be currency annotations (e.g. [$$-409])
				// color formats (e.g. [color1] through [color56], as well as [red] etc.)
				// conditionals (e.g. [>100], the valid conditionals are =, >, <, >=, <=, <>)
				bracketIndex, err := skipToRune(curReducedFormat, ']')
				if err != nil {
					// This is not any type of valid format.
					return false
				}
				i += bracketIndex
				continue
			}
			// Symbols that don't have meaning, aren't in the exempt literal characters, and aren't escaped are invalid.
			// The string could still be a valid number format string.
			return false
		}
	}
	// If the string doesn't have any time formatting characters, it could technically be a time format, but it
	// would be a pretty weak time format. A valid time format with no time formatting symbols will also be a number
	// format with no number formatting symbols, which is essentially a constant string that does not depend on the
	// cell's value in anyway. The downstream logic will do the right thing in that case if this returns false.
	return foundTimeFormatCharacters
}

func skipToRune(runes []rune, r rune) (int, error) {
	for i := 1; i < len(runes); i++ {
		if runes[i] == r {
			return i, nil
		}
	}
	return -1, fmt.Errorf("no closing quote found")
}

var timeFormatCharacters = []string{"M", "D", "Y", "YY", "YYYY", "MM", "yyyy", "m", "d", "yy", "h", "m", "AM/PM", "A/P", "am/pm", "a/p", "r", "g", "e", "b1", "b2", "[hh]", "[h]", "[mm]", "[m]",
	"s.0000", "s.000", "s.00", "s.0", "s", "[ss].0000", "[ss].000", "[ss].00", "[ss].0", "[ss]", "[s].0000", "[s].000", "[s].00", "[s].0", "[s]", "上", "午", "下"}

func parseLiterals(format string) (string, string, bool, error) {
	var prefix string
	showPercent := false
	for i := 0; i < len(format); i++ {
		curReducedFormat := format[i:]
		switch curReducedFormat[0] {
		case '\\':
			// If there is a slash, skip the next character, and add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
				prefix += curReducedFormat[1:2]
			}
		case '_':
			// If there is an underscore, skip the next character, but don't add it to the prefix
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
			// Asterisks are used to repeat the next character to fill the full cell width.
			// There isn't really a cell size in this context, so this will be ignored.
		case '"':
			// If there is a quote skip to the next quote, and add the quoted characters to the prefix
			endQuoteIndex := strings.Index(curReducedFormat[1:], "\"")
			if endQuoteIndex == -1 {
				return "", "", false, errors.New("invalid formatting code, unmatched double quote")
			}
			prefix = prefix + curReducedFormat[1:endQuoteIndex+1]
			i += endQuoteIndex + 1
		case '%':
			showPercent = true
			prefix += "%"
		case '[':
			// Brackets can be currency annotations (e.g. [$$-409])
			// color formats (e.g. [color1] through [color56], as well as [red] etc.)
			// conditionals (e.g. [>100], the valid conditionals are =, >, <, >=, <=, <>)
			bracketIndex := strings.Index(curReducedFormat, "]")
			if bracketIndex == -1 {
				return "", "", false, errors.New("invalid formatting code, invalid brackets")
			}
			// Currencies in Excel are annotated with this format: [$<Currency String>-<Language Info>]
			// Currency String is something like $, ¥, €, or £
			// Language Info is three hexadecimal characters
			if len(curReducedFormat) > 2 && curReducedFormat[1] == '$' {
				dashIndex := strings.Index(curReducedFormat, "-")
				if dashIndex != -1 && dashIndex < bracketIndex {
					// Get the currency symbol, and skip to the end of the currency format
					prefix += curReducedFormat[2:dashIndex]
				} else {
					return "", "", false, errors.New("invalid formatting code, invalid currency annotation")
				}
			}
			i += bracketIndex
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ':
			// These symbols are allowed to be used as literal without escaping
			prefix += curReducedFormat[0:1]
		default:
			for _, special := range formattingCharacters {
				if strings.HasPrefix(curReducedFormat, special) {
					// This means we found the start of the actual number formatting portion, and should return.
					return prefix, format[i:], showPercent, nil
				}
			}
			// Symbols that don't have meaning and aren't in the exempt literal characters and are not escaped.
			return "", "", false, errors.New("invalid formatting code: unsupported or unescaped characters")
		}
	}
	return prefix, "", showPercent, nil
}

var formattingCharacters = []string{"0/", "#/", "?/", "E-", "E+", "e-", "e+", "0", "#", "?", ".", ",", "@", "*"}

func splitFormatAndSuffixFormat(format string) (string, string) {
	var i int
	for ; i < len(format); i++ {
		curReducedFormat := format[i:]
		var found bool
		for _, special := range formattingCharacters {
			if strings.HasPrefix(curReducedFormat, special) {
				// Skip ahead if the special character was longer than length 1
				i += len(special) - 1
				found = true
				break
			}
		}
		if !found {
			break
		}
	}
	suffixFormat := format[i:]
	format = format[:i]
	return format, suffixFormat
}
