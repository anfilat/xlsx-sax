package xlsx

import (
	"bytes"
	"io"
	"os"
	"strconv"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNew(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)
	require.Len(t, xlsx.sheetNameFile, 2)
	require.Len(t, xlsx.sheetFile, 2)
	require.Len(t, xlsx.sharedStrings, 9)
}

func TestSheetNames(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	names := xlsx.SheetNames()
	require.Len(t, names, 2)
	require.Equal(t, "Sheet1", names[0])
	require.Equal(t, "Sheet2", names[1])
}

func TestOpenSheet(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(0)
	require.NoError(t, err)
	defer sheet.Close()

	err = sheet.SkipRow()
	require.NoError(t, err)

	isRow := sheet.NextRow()
	require.True(t, isRow)

	isCell := sheet.NextCell()
	require.True(t, isCell)
	val, err := sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 1, sheet.Row)
	require.Equal(t, 0, sheet.Col)
	require.Equal(t, "This is text, rich text", val)

	isCell = sheet.NextCell()
	require.True(t, isCell)
	val, err = sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 1, sheet.Row)
	require.Equal(t, 1, sheet.Col)
	require.Equal(t, "1245237", val)

	isCell = sheet.NextCell()
	require.True(t, isCell)
	val, err = sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 1, sheet.Row)
	require.Equal(t, 2, sheet.Col)
	require.Equal(t, "something", val)

	isRow = sheet.NextRow()
	require.True(t, isRow)

	isCell = sheet.NextCell()
	require.True(t, isCell)
	val, err = sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 2, sheet.Row)
	require.Equal(t, 0, sheet.Col)
	require.Equal(t, "The same", val)

	isCell = sheet.NextCell()
	require.True(t, isCell)
	val, err = sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 2, sheet.Row)
	require.Equal(t, 1, sheet.Col)
	require.Equal(t, "4534567", val)

	isCell = sheet.NextCell()
	require.True(t, isCell)
	val, err = sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, 2, sheet.Row)
	require.Equal(t, 2, sheet.Col)
	require.Equal(t, "a table", val)

	err = sheet.Err()
	require.NoError(t, err)

	err = sheet.Close()
	require.NoError(t, err)
}

func TestOpenEmptySheet(t *testing.T) {
	data, err := os.ReadFile("testdata/empty.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(0)
	require.NoError(t, err)
	defer sheet.Close()

	isRow := sheet.NextRow()
	require.False(t, isRow)

	err = sheet.Err()
	require.ErrorIs(t, err, io.EOF)
}

func TestOpenRichText(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(1)
	require.NoError(t, err)
	defer sheet.Close()

	isRow := sheet.NextRow()
	require.True(t, isRow)

	isCell := sheet.NextCell()
	require.True(t, isCell)
	val, err := sheet.CellValue()
	require.NoError(t, err)
	require.Equal(t, "This is text, rich text", val)
}

func TestReadSheet(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(0)
	require.NoError(t, err)
	defer sheet.Close()

	err = sheet.SkipRow()
	require.NoError(t, err)

	sum := 0
	for sheet.NextRow() {
		for sheet.NextCell() {
			if sheet.Col == 3 {
				val, er := sheet.CellValue()
				require.NoError(t, er)

				n, er := strconv.Atoi(val)
				require.NoError(t, er)
				sum += n
			}
		}
	}
	err = sheet.Err()
	require.ErrorIs(t, err, io.EOF)
	require.Equal(t, 5, sum)
}

func TestMultiRow(t *testing.T) {
	// Read the file
	data, err := os.ReadFile("testdata/multi_row.xlsx")
	require.NoError(t, err)

	// Create a new XLSX reader
	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	// Open the first sheet
	sheet, err := xlsx.OpenSheetByOrder(0)
	require.NoError(t, err)
	defer sheet.Close()

	// Check first row
	isRow := sheet.NextRow()
	require.True(t, isRow)

	expectedFirstRow := []string{"Text", "Id", "Name", "Count", "Fill"}
	for i, expected := range expectedFirstRow {
		isCell := sheet.NextCell()
		require.True(t, isCell)
		val, err := sheet.CellValue()
		require.NoError(t, err)
		require.Equal(t, expected, val)
		require.Equal(t, i, sheet.Col)
	}

	// Check second row
	isRow = sheet.NextRow()
	require.True(t, isRow)

	expectedSecondRow := []string{"Some text", "1245237", "something", "5", "Filled"}
	for i, expected := range expectedSecondRow {
		isCell := sheet.NextCell()
		require.True(t, isCell)
		val, err := sheet.CellValue()
		require.NoError(t, err)
		require.Equal(t, expected, val)
		require.Equal(t, i, sheet.Col)
	}

	// Verify there are no more rows
	isRow = sheet.NextRow()
	require.False(t, isRow)

	// Check for errors
	err = sheet.Err()
	require.ErrorIs(t, err, io.EOF)
}

func BenchmarkXlsx1(b *testing.B) {
	data, _ := os.ReadFile("testdata/test1.xlsx")
	br := bytes.NewReader(data)

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		xlsx, _ := New(br, br.Size())
		sheet, _ := xlsx.OpenSheetByOrder(0)

		for sheet.NextRow() {
			item := xlsx1Item{}
			for sheet.NextCell() {
				if sheet.Col == 0 {
					item.Name, _ = sheet.CellValue()
				} else if sheet.Col == 1 {
					item.Offer, _ = sheet.CellFormatValue()
				} else if sheet.Col == 3 {
					item.Count, _ = sheet.CellInt()
				}
			}
		}

		_ = sheet.Close()
	}
}

type xlsx1Item struct {
	Name  string
	Offer string
	Count int
}
