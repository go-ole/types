// +build windows

package iunknown

import (
	"unsafe"

	"github.com/jacobsantos/go-com/com"
	syscall "golang.org/x/sys/windows"
)

// QueryInterfaceService implements QueryInterface for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func QueryInterfaceService(self *interface{}, interfaceID *com.GUID, client **interface{}) uint32 {
	obj := (*com.ServiceReference)(unsafe.Pointer(self))
	*client = nil
	if interfaceID.IsEqual(InterfaceID) || interfaceID.IsEqual(obj.InterfaceID) {
		AddRefCallback(self)
		*client = self
		return com.SuccessResponseCode
	}
	return com.NoInterfaceErrorCode
}

// AddRefService implements AddRef for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func AddRefService(self *interface{}) int32 {
	obj := (*UnknownServer)(unsafe.Pointer(self))
	obj.ReferenceCount++
	return obj.ReferenceCount
}

// ReleaseService implements Release for IUnknown interface.
//
// This is for COM server or IUnknown callback.
func ReleaseService(self *interface{}) int32 {
	obj := (*UnknownServer)(unsafe.Pointer(self))
	obj.ReferenceCount--
	return obj.ReferenceCount
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
