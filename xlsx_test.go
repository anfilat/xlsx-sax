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
	require.Len(t, xlsx.sharedStrings, 1)
}

func TestOpenSheet(t *testing.T) {
	data, err := os.ReadFile("testdata/test1.xlsx")
	require.NoError(t, err)

	br := bytes.NewReader(data)
	xlsx, err := New(br, br.Size())
	require.NoError(t, err)

	_, err = xlsx.OpenSheetByOrder(0, &SheetParams{
		Skip: 1,
	})
	require.NoError(t, err)
}
