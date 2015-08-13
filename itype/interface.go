package itype

import (
	"unsafe"

	"github.com/go-ole/iunknown"
	syscall "golang.org/x/sys/windows"
)

type Info iunknown.Unknown

type Compiler iunknown.Unknown

type Library iunknown.Unknown

type InfoVirtualTable struct {
	iunknown.VirtualTable
	GetTypeAttr          uintptr
	GetTypeComp          uintptr
	GetFuncDesc          uintptr
	GetVarDesc           uintptr
	GetNames             uintptr
	GetRefTypeOfImplType uintptr
	GetImplTypeFlags     uintptr
	GetIDsOfNames        uintptr
	Invoke               uintptr
	GetDocumentation     uintptr
	GetDllEntry          uintptr
	GetRefTypeInfo       uintptr
	AddressOfMember      uintptr
	CreateInstance       uintptr
	GetMops              uintptr
	GetContainingTypeLib uintptr
	ReleaseTypeAttr      uintptr
	ReleaseFuncDesc      uintptr
	ReleaseVarDesc       uintptr
}

func (v *Info) VTable() *InfoVirtualTable {
	return (*InfoVirtualTable)(unsafe.Pointer(v.RawVTable))
}

func (v *Info) GetTypeAttr() (tattr *TYPEATTR, err error) {
	hr, _, _ := syscall.Syscall(
		uintptr(v.VTable().GetTypeAttr),
		2,
		uintptr(unsafe.Pointer(v)),
		uintptr(unsafe.Pointer(&tattr)),
		0)
	if hr != 0 {
		err = NewError(hr)
	}
	return
}
