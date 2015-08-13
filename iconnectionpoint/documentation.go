// +build !windows

package iconnectionpoint

import (
	"github.com/go-ole/com"
	"github.com/go-ole/iunknown"
)

func Advise(obj interface{}, method uintptr, unknown *iunknown.IUnknown) (uint32, error) {
	return uint32(0), com.NotImplementedError
}

func Unadvise(obj interface{}, method uintptr, cookie uint32) error {
	return com.NotImplementedError
}

func GetConnectionInterface(obj interface{}, method uintptr) (*com.GUID, error) {
	return nil, com.NotImplementedError
}

func GetConnectionPointContainer(obj interface{}, method uintptr) (*ConnectionPointContainer, error) {
	return nil, com.NotImplementedError
}

// EnumConnections creates an enumerator object to iterate through current
// connections.
//
// XXX: Need IEnumConnections structure
func EnumConnections(obj interface{}, method uintptr) (interface{}, error) {
	return nil, com.NotImplementedError
}

// EnumConnectionPoints creates an enumerator object to iterate through
// connection points.
//
// XXX: Need to implement IEnumConnectionPoints structure.
func EnumConnectionPoints(obj interface{}, method uintptr) (interface{}, error) {
	return nil, com.NotImplementedError
}

func FindConnectionPoint(obj interface{}, method uintptr, interfaceID *com.GUID) (*ConnectionPoint, error) {
	return nil, com.NotImplementedError
}
