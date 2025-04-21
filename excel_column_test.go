package xlsx

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestColumnIndex(t *testing.T) {
	require.Equal(t, 0, columnIndex([]byte("A")))
	require.Equal(t, 1, columnIndex([]byte("B")))
	require.Equal(t, 25, columnIndex([]byte("Z")))
	require.Equal(t, 26, columnIndex([]byte("AA")))
	require.Equal(t, 27, columnIndex([]byte("AB")))
	require.Equal(t, 27, columnIndex([]byte("AB33")))
}
