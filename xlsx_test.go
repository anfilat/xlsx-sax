package xlsx

import (
	"bytes"
	"fmt"
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

func BenchmarkXlsx1(b *testing.B) {
	data, _ := os.ReadFile("testdata/test1.xlsx")
	br := bytes.NewReader(data)

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		xlsx, _ := New(br, br.Size())
		sheet, _ := xlsx.OpenSheetByOrder(0)

		for sheet.NextRow() {
			for sheet.NextCell() {
				val, _ := sheet.CellValue()
				if len(val) > 100000 {
					fmt.Println(val)
				}
			}
		}

		_ = sheet.Close()
	}
}
