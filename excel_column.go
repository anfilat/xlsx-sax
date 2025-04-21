package xlsx

func columnIndex(s string) int {
	result := 0
	ar := []rune(s)

	for i, j := len(ar)-2, 26; i >= 0; i, j = i-1, j*26 {
		result += (int(ar[i]-'A') + 1) * j
	}

	return result + int(ar[len(ar)-1]-'A')
}
