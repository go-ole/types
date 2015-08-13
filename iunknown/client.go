// +build windows

package iunknown

import (
	"reflect"
	"unsafe"

	"github.com/jacobsantos/go-com/com"
	syscall "golang.org/x/sys/windows"
)

// QueryInterface returns object matching InterfaceID if it exists.
//
// Method must match QueryInterface function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
//
// Client must be passed by reference.
func QueryInterface(obj interface{}, method uintptr, interfaceID *com.GUID, client interface{}) error {
	return com.HResultToError(syscall.Syscall(
		method,
		uintptr(3),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(interfaceID)),
		reflect.ValueOf(client).UnsafeAddr()))
}

// AddRef increments the reference counter for object.
//
// Method must match AddRef function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
func AddRef(obj interface{}, method uintptr) int32 {
	return com.GetInt32FromCall(uintptr(unsafe.Pointer(obj)), method)
}

// Release decrements the reference counter for object.
//
// Method must match Release function pointer or the the behavior will be
// undefined. Resulting possibly in a crash or incorrect results.
func Release(obj interface{}, method uintptr) int32 {
	return com.GetInt32FromCall(uintptr(unsafe.Pointer(obj)), method)
}
