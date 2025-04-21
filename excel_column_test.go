package xlsx

import (
	"github.com/stretchr/testify/require"
	"testing"
)

func TestColumnIndex(t *testing.T) {
	require.Equal(t, 0, columnIndex("A"))
	require.Equal(t, 1, columnIndex("B"))
	require.Equal(t, 25, columnIndex("Z"))
	require.Equal(t, 26, columnIndex("AA"))
	require.Equal(t, 27, columnIndex("AB"))
	require.Equal(t, 27, columnIndex("AB33"))
}
