// +build windows

package iconnectionpoint

import (
	"reflect"
	"unsafe"

	syscall "golang.org/x/sys/windows"

	"github.com/go-ole/com"
	"github.com/go-ole/iunknown"
)

func Advise(obj interface{}, method uintptr, unknown *iunknown.IUnknown) (cookie uint32, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(3),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(unknown)),
		uintptr(unsafe.Pointer(&cookie))))
	return
}

func Unadvise(obj interface{}, method uintptr, cookie uint32) error {
	return com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		uintptr(cookie),
		uintptr(0)))
}

func GetConnectionInterface(obj interface{}, method uintptr) (interfaceID *com.GUID, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&interfaceID)),
		uintptr(0)))
	return
}

func GetConnectionPointContainer(obj interface{}, method uintptr) (element *iconnectionpointcontainer.ConnectionPointContainer, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&element)),
		uintptr(0)))
	return 0
}

// EnumConnections creates an enumerator object to iterate through current
// connections.
//
// XXX: Need IEnumConnections structure
func EnumConnections(obj interface{}, method uintptr) (element *interface{}, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&element)),
		uintptr(0)))
	return 0
}

// EnumConnectionPoints creates an enumerator object to iterate through
// connection points.
//
// XXX: Need to implement IEnumConnectionPoints structure.
func EnumConnectionPoints(obj interface{}, method uintptr) (element interface{}, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		reflect.ValueOf(element).UnsafeAddr(),
		uintptr(0)))
	return
}

func FindConnectionPoint(obj interface{}, method uintptr, interfaceID *com.GUID) (element *iconnectionpoint.ConnectionPoint, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(3),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(interfaceID)),
		uintptr(unsafe.Pointer(&element))))
	return
}
