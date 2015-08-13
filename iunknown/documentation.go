// +build !windows

package iunknown

import "syscall"

import "github.com/go-ole/com"

const zero = int32(0)

// QueryInterface returns object matching InterfaceID if it exists.
//
// Method must match QueryInterface function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
//
// Client must be passed by reference.
func QueryInterface(obj interface{}, method uintptr, interfaceID *com.GUID, client interface{}) error {
	return com.NotImplementedError
}

// AddRef increments the reference counter for object.
//
// Method must match AddRef function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
func AddRef(obj interface{}, method uintptr) int32 {
	return zero
}

// Release decrements the reference counter for object.
//
// Method must match Release function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
func Release(obj interface{}, method uintptr) int32 {
	return zero
}

// QueryInterfaceService implements QueryInterface for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func QueryInterfaceService(self *interface{}, interfaceID *com.GUID, client **interface{}) uint32 {
	return com.NoInterfaceErrorCode
}

// AddRefService implements AddRef for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func AddRefService(self *interface{}) int32 {
	return zero
}

// ReleaseService implements Release for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func ReleaseService(self *interface{}) int32 {
	return zero
}

// VirtualTableService creates VirtualTable with references to IUnknown methods.
//
// This should be used with Service().
func VirtualTableService() *VirtualTable {
	return &VirtualTable{
		QueryInterface: syscall.NewCallback(QueryInterfaceService),
		AddRef:         syscall.NewCallback(AddRefService),
		Release:        syscall.NewCallback(ReleaseService)}
}

// IUnknownService creates IUnknown interface COM server reference object.
func IUnknownService(obj *interface{}) (*com.ServerObject, error) {
	return Service(obj, VirtualTableService(), InterfaceID)
}
