package iunknown

import (
	"github.com/jacobsantos/go-com/idispatch"
)

func ExampleUnknown_QueryInterface_idispatch() {
	var unknown Unknown
	var dispatch idispatch.Dispatch
	unknown.QueryInterface(idispatch.InterfaceID, &dispatch)
	// Output:
}
