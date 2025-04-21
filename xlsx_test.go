package xlsx

import (
	"bytes"
	"os"
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
	require.Len(t, xlsx.sharedStrings, 7)
}

func TestOpenSheet(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(0, []bool{true, true, false, true}, 1)
	require.NoError(t, err)

	isRow := sheet.Next()
	require.True(t, isRow)

	row, err := sheet.Read()
	require.NoError(t, err)
	require.Equal(t, []string{"This is text, rich text", "1245237", "5"}, row)

	row, err = sheet.Read()
	require.NoError(t, err)
	require.Equal(t, []string{"The same", "4534567", "0"}, row)

	err = sheet.Close()
	require.NoError(t, err)
}

func TestOpenEmptySheet(t *testing.T) {
	data, err := os.ReadFile("testdata/empty.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	sheet, err := xlsx.OpenSheetByOrder(0, []bool{true, true, false, true}, 1)
	require.NoError(t, err)

	isRow := sheet.Next()
	require.False(t, isRow)
}
