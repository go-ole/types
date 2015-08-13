// +build !windows

package idispatch

import (
	"syscall"

	"github.com/go-ole/com"
	"github.com/go-ole/itypeinfo"
)

func GetIDsOfName(obj interface{}, method uintptr, names []string) ([]int32, error) {
	return []int32{}, com.NotImplementedError
}

func GetTypeInfoCount(obj interface{}, method uintptr) (uint32, error) {
	return uint32(0), com.NotImplementedError
}

func GetTypeInfo(obj interface{}, method uintptr, num uint32) (*itypeinfo.TypeInfo, error) {
	return nil, com.NotImplementedError
}

func Invoke(obj interface{}, method uintptr, displayID DisplayID, dispatchContext DispatchContext, params ...interface{}) (*com.Variant, error) {
	return nil, com.NotImplementedError
}

func GetIDsOfNamesService(this *iunknown.IUnknown, iid *com.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
	return com.NotImplementedErrorCode
}

func GetTypeInfoCountService(pcount *int) uintptr {
	return com.NotImplementedErrorCode
}

func GetTypeInfoService(ptypeif *uintptr) uintptr {
	return com.NotImplementedErrorCode
}

func InvokeService(this *IDispatch, dispid int32, riid *com.GUID, lcid int, flags int16, dispparams *com.DISPPARAMS, result *com.Variant, pexcepinfo *com.EXCEPINFO, nerr *uint) uintptr {
	return com.NotImplementedErrorCode
}

// VirtualTableService creates VirtualTable with references to IUnknown methods.
//
// This should be used with Service().
func VirtualTableService() *VirtualTable {
	return &VirtualTable{
		QueryInterface:   syscall.NewCallback(iunknown.QueryInterfaceService),
		AddRef:           syscall.NewCallback(iunknown.AddRefService),
		Release:          syscall.NewCallback(iunknown.ReleaseService),
		GetTypeInfoCount: syscall.NewCallback(GetTypeInfoCountService),
		GetTypeInfo:      syscall.NewCallback(GetTypeInfoService),
		GetIDsOfNames:    syscall.NewCallback(GetIDsOfNamesService),
		Invoke:           syscall.NewCallback(InvokeService)}
}

// IDispatchService creates IUnknown interface COM server reference object.
func IDispatchService(obj *interface{}) (*com.ServerObject, error) {
	return Service(obj, VirtualTableService(), InterfaceID)
}
