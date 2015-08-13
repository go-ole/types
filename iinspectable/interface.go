// +build windows
package iinspectable

import (
	"reflect"
	"syscall"
	"unsafe"

	"github.com/jacobsantos/go-com/pkg/com"
	"github.com/jacobsantos/go-com/pkg/iunknown"
)

// IID for IInspectable
var InterfaceID = &com.GUID{0xaf86e2e0, 0xb12d, 0x4c6a, [8]byte{0x9c, 0x5a, 0xd7, 0xaa, 0x65, 0x10, 0x1e, 0x90}}

// Inspectable is the structure for storing raw virtual table.
type Inspectable iunknown.Unknown

// The Virtual Table for IInspectable COM objects.
//
// Stores pointers to IInspectable methods to be called using Syscall or unsafe.
type VirtualTable struct {
	iunknown.VirtualTable
	GetInterfaceIDs     uintptr
	GetRuntimeClassName uintptr
	GetTrustLevel       uintptr
}

type IInspectable interface {
	GetInterfaceIDs()
	GetRuntimeClassName()
	GetTrustLevel()
}

func (i *Inspectable) VTable() *VirtualTable {
	return (*VirtualTable)(unsafe.Pointer(i.RawVTable))
}

// QueryInterface will query interface and return object by reference.
func (i *Inspectable) QueryInterface(interfaceID *com.GUID, element interface{}) error {
	return iunknown.QueryInterface(i, i.VTable().QueryInterface, interfaceID, &element)
}

// AddRef increments the reference counter for object.
func (i *Inspectable) AddRef() int32 {
	return iunknown.AddRef(i, i.VTable().AddRef)
}

// Release decrements the reference counter for object.
func (i *Inspectable) Release() int32 {
	return iunknown.Release(i, i.VTable().Release)
}

func (v *Inspectable) GetIids() (iids []*GUID, err error) {
	var count uint32
	var array uintptr
	hr, _, _ := syscall.Syscall(
		v.VTable().GetIIds,
		3,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&count)),
		uintptr(unsafe.Pointer(&array)))
	if hr != 0 {
		err = NewError(hr)
		return
	}
	defer CoTaskMemFree(array)

	iids = com.PointerToArray(array, count, reflect.TypeOf(*GUID)).([]*GUID)
	return
}

func (v *Inspectable) GetRuntimeClassName() (s string, err error) {
	var hstring HString
	err = HResultToError(syscall.Syscall(
		v.VTable().GetRuntimeClassName,
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&hstring)),
		0))

	if err != nil {
		return
	}

	s = hstring.String()
	DeleteHString(hstring)
	return
}

func (v *Inspectable) GetTrustLevel() (level uint32, err error) {
	err = HResultToError(syscall.Syscall(
		v.VTable().GetTrustLevel,
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&level)),
		0))
	return
}
