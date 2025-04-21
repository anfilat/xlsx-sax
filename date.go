package xlsx

// Most of this file was taken from https://github.com/tealeg/xlsx

import "time"

const nanosInADay = float64((24 * time.Hour) / time.Nanosecond)

var (
	excel1900Epoc = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	excel1904Epoc = time.Date(1904, time.January, 1, 0, 0, 0, 0, time.UTC)
)

func timeFromExcelTime(excelTime float64, date1904 bool) time.Time {
	wholeDaysPart := int(excelTime)
	durationPart := time.Duration(nanosInADay * (excelTime - float64(wholeDaysPart)))
	if date1904 {
		return excel1904Epoc.AddDate(0, 0, wholeDaysPart).Add(durationPart)
	}
	return excel1900Epoc.AddDate(0, 0, wholeDaysPart).Add(durationPart)
}
