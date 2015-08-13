// +build windows
package iprovideclassinfo

import (
	"unsafe"

	"github.com/jacobsantos/go-com/pkg/com"
	"github.com/jacobsantos/go-com/pkg/itypeinfo"
	"github.com/jacobsantos/go-com/pkg/iunknown"
)

// IID for IUnknown
var InterfaceID = &com.GUID{0xb196b283, 0xbab4, 0x101a, [8]byte{0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07}}

type ProvideClassInfo iunknown.Unknown

// The Virtual Table for ITypeLib COM objects.
//
// Stores pointers to ITypeLib methods to be called using Syscall or unsafe.
type VirtualTable struct {
	iunknown.VirtualTable
	GetClassInfo uintptr
}

type IProvideClassInfo interface {
	GetClassInfo() (*itypeinfo.TypeInfo, error)
}

func (p *ProvideClassInfo) VTable() *VirtualTable {
	return (*VirtualTable)(unsafe.Pointer(p.RawVTable))
}

// QueryInterface will query interface and return object by reference.
func (p *ProvideClassInfo) QueryInterface(interfaceID *com.GUID, element interface{}) error {
	return iunknown.QueryInterface(p, p.VTable().QueryInterface, interfaceID, &element)
}

// AddRef increments the reference counter for object.
func (p *ProvideClassInfo) AddRef() int32 {
	return iunknown.AddRef(p, p.VTable().AddRef)
}

// Release decrements the reference counter for object.
func (p *ProvideClassInfo) Release() int32 {
	return iunknown.Release(p, p.VTable().Release)
}

func (p *ProvideClassInfo) GetClassInfo() (*itypeinfo.TypeInfo, error) {
	return GetClassInfo(p, p.VTable().GetClassInfo)
}
