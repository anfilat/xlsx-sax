package xlsx

// Most of this file was taken from https://github.com/tealeg/xlsx

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

const builtinNumFormatsCount = 163

var builtinNumFormats = []string{
	0:  "general",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00e+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm am/pm",
	19: "h:mm:ss am/pm",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[red](#,##0.00)",
	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0e+0",
	49: "@",
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
		parsedNumFmt.isTimeFormat = fmtOptions[0].isTimeFormat
		parsedNumFmt.negativeFormatExpectsPositive = true
		parsedNumFmt.positiveFormat = fmtOptions[0]
		parsedNumFmt.negativeFormat = fmtOptions[1]
		parsedNumFmt.zeroFormat = fmtOptions[0]
		parsedNumFmt.textFormat, _ = parseNumberFormatSection("general")
	} else if len(fmtOptions) == 3 {
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

	if compareFormatString(reducedFormat, "general") {
		return &formatOptions{
			fullFormatString:    "general",
			reducedFormatString: "general",
		}, nil
	}

	if isTimeFormat(reducedFormat) {
		return &formatOptions{
			isTimeFormat:        true,
			fullFormatString:    fullFormat,
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
		return nil, ErrInvalidFormat
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
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
		case '"':
			endQuoteIndex, err := skipToRune(curReducedFormat, '"')
			if err != nil {
				return false
			}
			i += endQuoteIndex + 1
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ':
		case ',':
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
				bracketIndex, err := skipToRune(curReducedFormat, ']')
				if err != nil {
					return false
				}
				i += bracketIndex
				continue
			}
			return false
		}
	}
	return foundTimeFormatCharacters
}

func skipToRune(runes []rune, r rune) (int, error) {
	for i := 1; i < len(runes); i++ {
		if runes[i] == r {
			return i, nil
		}
	}
	return -1, ErrNoClosingQuote
}

var timeFormatCharacters = []string{
	"M", "D", "Y", "YY", "YYYY", "MM", "yyyy", "m", "d", "yy", "h", "m", "AM/PM", "A/P", "am/pm", "a/p", "r", "g", "e", "b1", "b2", "[hh]", "[h]", "[mm]", "[m]",
	"s.0000", "s.000", "s.00", "s.0", "s", "[ss].0000", "[ss].000", "[ss].00", "[ss].0", "[ss]", "[s].0000", "[s].000", "[s].00", "[s].0", "[s]", "上", "午", "下",
}

func parseLiterals(format string) (string, string, bool, error) {
	var prefix string
	showPercent := false
	for i := 0; i < len(format); i++ {
		curReducedFormat := format[i:]
		switch curReducedFormat[0] {
		case '\\':
			if len(curReducedFormat) > 1 {
				i++
				prefix += curReducedFormat[1:2]
			}
		case '_':
			if len(curReducedFormat) > 1 {
				i++
			}
		case '*':
		case '"':
			endQuoteIndex := strings.Index(curReducedFormat[1:], "\"")
			if endQuoteIndex == -1 {
				return "", "", false, ErrDoubleQuote
			}
			prefix = prefix + curReducedFormat[1:endQuoteIndex+1]
			i += endQuoteIndex + 1
		case '%':
			showPercent = true
			prefix += "%"
		case '[':
			bracketIndex := strings.Index(curReducedFormat, "]")
			if bracketIndex == -1 {
				return "", "", false, ErrInvalidBrackets
			}
			if len(curReducedFormat) > 2 && curReducedFormat[1] == '$' {
				dashIndex := strings.Index(curReducedFormat, "-")
				if dashIndex != -1 && dashIndex < bracketIndex {
					prefix += curReducedFormat[2:dashIndex]
				} else {
					return "", "", false, ErrInvalidCurrency
				}
			}
			if curReducedFormat[1] == '=' || curReducedFormat[1] == '>' || curReducedFormat[1] == '<' {
				return "", "", false, fmt.Errorf("unsupported formatting code: %s", format)
			}
			i += bracketIndex
		case '$', '-', '+', '/', '(', ')', ':', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=', ' ':
			prefix += curReducedFormat[0:1]
		default:
			for _, special := range formattingCharacters {
				if strings.HasPrefix(curReducedFormat, special) {
					return prefix, format[i:], showPercent, nil
				}
			}
			return "", "", false, EUnsupportedCharacters
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

func (p *parsedNumFormat) text(value string) (string, error) {
	textFormat := p.textFormat
	switch textFormat.reducedFormatString {
	case "general":
		return value, nil
	case "@":
		return textFormat.prefix + value + textFormat.suffix, nil
	case "":
		return textFormat.prefix + textFormat.suffix, nil
	default:
		return value, ErrInvalidFormat
	}
}

func (p *parsedNumFormat) numeric(value string, date1904 bool) (string, error) {
	rawValue := strings.TrimSpace(value)
	if rawValue == "" {
		return "", nil
	}

	if p.isTimeFormat {
		return p.parseTime(rawValue, date1904)
	}
	var numberFormat *formatOptions
	floatVal, floatErr := strconv.ParseFloat(rawValue, 64)
	if floatErr != nil {
		return rawValue, floatErr
	}

	if floatVal > 0 {
		numberFormat = p.positiveFormat
	} else if floatVal < 0 {
		if p.negativeFormatExpectsPositive {
			floatVal = math.Abs(floatVal)
		}
		numberFormat = p.negativeFormat
	} else {
		numberFormat = p.zeroFormat
	}

	if numberFormat.showPercent {
		floatVal = 100 * floatVal
	}

	var formattedNum string
	switch numberFormat.reducedFormatString {
	case "general":
		generalFormatted, err := generalNumericScientific(value, true)
		if err != nil {
			return rawValue, nil
		}
		return generalFormatted, nil
	case "@":
		formattedNum = value
	case "0", "#,##0":
		formattedNum = fmt.Sprintf("%.0f", floatVal)
	case "0.0", "#,##0.0":
		formattedNum = fmt.Sprintf("%.1f", floatVal)
	case "0.00", "#,##0.00":
		formattedNum = fmt.Sprintf("%.2f", floatVal)
	case "0.000", "#,##0.000":
		formattedNum = fmt.Sprintf("%.3f", floatVal)
	case "0.0000", "#,##0.0000":
		formattedNum = fmt.Sprintf("%.4f", floatVal)
	case "0.00e+00", "##0.0e+0":
		formattedNum = fmt.Sprintf("%e", floatVal)
	case "":
		//
	default:
		if cntZeros := strings.Count(numberFormat.reducedFormatString, "0"); cntZeros == len(numberFormat.reducedFormatString) && cntZeros > len(rawValue) {
			return numberFormat.reducedFormatString[:cntZeros-len(rawValue)] + rawValue, nil
		}
		return rawValue, nil
	}
	return numberFormat.prefix + formattedNum + numberFormat.suffix, nil
}

var timeReplacements = []struct{ xltime, gotime string }{
	{"YYYY", "2006"},
	{"yyyy", "2006"},
	{"YY", "06"},
	{"yy", "06"},
	{"MMMM", "%%%%"},
	{"mmmm", "%%%%"},
	{"DDDD", "&&&&"},
	{"dddd", "&&&&"},
	{"DD", "02"},
	{"dd", "02"},
	{"D", "2"},
	{"d", "2"},
	{"MMM", "Jan"},
	{"mmm", "Jan"},
	{"MMSS", "0405"},
	{"mmss", "0405"},
	{"SS", "05"},
	{"ss", "05"},
	{"MM:", "04:"},
	{"mm:", "04:"},
	{":MM", ":04"},
	{":mm", ":04"},
	{"MM", "01"},
	{"mm", "01"},
	{"AM/PM", "pm"},
	{"am/pm", "pm"},
	{"M/", "1/"},
	{"m/", "1/"},
	{"%%%%", "January"},
	{"&&&&", "Monday"},
}

func (p *parsedNumFormat) parseTime(value string, date1904 bool) (string, error) {
	f, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value, err
	}
	val := timeFromExcelTime(f, date1904)
	format := p.positiveFormat.fullFormatString
	if is12HourTime(format) {
		format = strings.Replace(format, "hh", "03", 1)
		format = strings.Replace(format, "h", "3", 1)
	} else {
		format = strings.Replace(format, "hh", "15", 1)
		format = strings.Replace(format, "h", "15", 1)
	}
	for _, repl := range timeReplacements {
		format = strings.Replace(format, repl.xltime, repl.gotime, 1)
	}
	if val.Hour() < 1 {
		format = strings.Replace(format, "]:", "]", 1)
		format = strings.Replace(format, "[03]", "", 1)
		format = strings.Replace(format, "[3]", "", 1)
		format = strings.Replace(format, "[15]", "", 1)
	} else {
		format = strings.Replace(format, "[3]", "3", 1)
		format = strings.Replace(format, "[15]", "15", 1)
	}
	return val.Format(format), nil
}

const (
	maxNonScientificNumber = 1e11
	minNonScientificNumber = 1e-9
)

func generalNumericScientific(value string, allowScientific bool) (string, error) {
	if strings.TrimSpace(value) == "" {
		return "", nil
	}
	f, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return value, err
	}
	if allowScientific {
		absF := math.Abs(f)
		if (absF >= math.SmallestNonzeroFloat64 && absF < minNonScientificNumber) || absF >= maxNonScientificNumber {
			return strconv.FormatFloat(f, 'E', -1, 64), nil
		}
	}
	return strconv.FormatFloat(f, 'f', -1, 64), nil
}

func is12HourTime(format string) bool {
	return strings.Contains(format, "am/pm") || strings.Contains(format, "AM/PM") || strings.Contains(format, "a/p") || strings.Contains(format, "A/P")
}
