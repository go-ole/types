// +build windows

package idispatch

import (
	"reflect"
	"unsafe"

	syscall "golang.org/x/sys/windows"

	"github.com/go-ole/com"
	"github.com/go-ole/iunknown"
)

func GetIDsOfNamesService(this *iunknown.IUnknown, iid *com.GUID, wnames []*uint16, namelen int, lcid int, pdisp []int32) uintptr {
	pthis := (*com.ServerObject)(unsafe.Pointer(this))
	names := make([]string, len(wnames))
	for i := 0; i < len(names); i++ {
		names[i] = com.LpOleStrToString(wnames[i])
	}
	for n := 0; n < namelen; n++ {
		if id, ok := pthis.funcMap[names[n]]; ok {
			pdisp[n] = id
		}
	}
	return com.SuccessResponseCode
}

func GetTypeInfoCountService(pcount *int) uintptr {
	if pcount != nil {
		*pcount = 0
	}
	return com.SuccessResponseCode
}

func GetTypeInfoService(ptypeif *uintptr) uintptr {
	return com.NotImplementedErrorCode
}

func InvokeService(this *IDispatch, dispid int32, riid *com.GUID, lcid int, flags int16, dispparams *com.DISPPARAMS, result *com.Variant, pexcepinfo *com.EXCEPINFO, nerr *uint) uintptr {
	pthis := (*com.ServerObject)(unsafe.Pointer(this))
	found := ""
	for name, id := range pthis.funcMap {
		if id == dispid {
			found = name
		}
	}
	if found != "" {
		rv := reflect.ValueOf(pthis.iface).Elem()
		rm := rv.MethodByName(found)
		rr := rm.Call([]reflect.Value{})
		println(len(rr))
		return com.SuccessResponseCode
	}
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
