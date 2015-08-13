package iunknown

import (
	"unsafe"

	"github.com/jacobsantos/go-com/com"
)

// InterfaceID is the IID for IUnknown.
var InterfaceID = &com.GUID{0x00000000, 0x0000, 0x0000, [8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

// Unknown is the structure for storing raw virtual table.
type Unknown struct {
	RawVTable *interface{}
}

// VirtualTable is the Virtual Table for IUnknown COM objects.
//
// Stores pointers to IUnknown methods to be called using Syscall or unsafe.
type VirtualTable struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

// IUnknown interface allows for typing Unknown objects.
type IUnknown interface {
	QueryInterface(interfaceID *com.GUID, element interface{}) error
	AddRef() uint32
	Release() uint32
}

// VTable converts Unknown to IUnknown VirtualTable.
//
// This is really an internal method for use with syscall. It is public in case
// it is required for future use and other structures will override.
func (u *Unknown) VTable() *VirtualTable {
	return (*VirtualTable)(unsafe.Pointer(u.RawVTable))
}

// QueryInterface will query interface and return object by reference.
func (u *Unknown) QueryInterface(interfaceID *com.GUID, element interface{}) error {
	return QueryInterface(u, u.VTable().QueryInterface, interfaceID, &element)
}

// MustQueryInterface will query interface and panic on failure.
func (u *Unknown) MustQueryInterface(interfaceID *com.GUID, element interface{}) {
	err = QueryInterface(u, u.VTable().QueryInterface, interfaceID, &element)
	if err != nil {
		panic(err)
	}
	return
}

// AddRef increments the reference counter for object.
func (u *Unknown) AddRef() int32 {
	return AddRef(u, u.VTable().AddRef)
}

// Release decrements the reference counter for object.
func (u *Unknown) Release() int32 {
	return Release(u, u.VTable().Release)
}
