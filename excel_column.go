package xlsx

func columnIndex(s []byte) int {
	result := 0
	for _, r := range s {
		if r >= '0' && r <= '9' {
			break
		}
		result = result*26 + int(r-'A') + 1
	}
	return result - 1
}
