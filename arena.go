package xlsx

import "unsafe"

type arena struct {
	alloc []byte
}

func (a *arena) toString(b []byte) string {
	n := len(b)
	if cap(a.alloc)-len(a.alloc) < n {
		a.reserve(n)
	}

	pos := len(a.alloc)
	data := a.alloc[pos : pos+n : pos+n]
	a.alloc = a.alloc[:pos+n]

	copy(data, b)

	return *(*string)(unsafe.Pointer(&data))
}

func (a *arena) reserve(n int) {
	a.alloc = make([]byte, 0, max(16*1024, n))
}
