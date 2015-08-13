// +build windows

package iunknown

import (
	"github.com/jacobsantos/go-com/pkg/idispatch"
)

func ExampleQueryInterface_idispatch() {
	// Unknown must be initialized through CoCreateInterface.
	var unknown Unknown
	var dispatch idispatch.Dispatch
	QueryInterface(unknown, unknown.VTable(), idispatch.InterfaceID, &dispatch)
	// Output:
}
