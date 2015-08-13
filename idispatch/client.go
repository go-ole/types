// +build windows

package idispatch

import (
	"unsafe"

	syscall "golang.org/x/sys/windows"

	"github.com/go-ole/com"
	"github.com/go-ole/itypeinfo"
)

func GetIDsOfName(obj interface{}, method uintptr, names []string) (displayIDs []int32, err error) {
	wnames := make([]*uint16, len(names))
	for i := 0; i < len(names); i++ {
		ptr, err := syscall.UTF16PtrFromString(names[i])
		if err != nil {
			return
		}
		wnames[i] = ptr
	}

	displayIDs = make([]int32, len(names))
	dispIDs := make([]int32, len(names))
	namelen := uint32(len(names))

	err = com.HResultToError(syscall.Syscall6(
		method,
		uintptr(6),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(com.NullInterfaceID)),
		uintptr(unsafe.Pointer(&wnames[0])),
		uintptr(namelen),
		uintptr(com.GetDefaultUserLocaleID()),
		uintptr(unsafe.Pointer(&dispIDs[0]))))

	displayIDs = dispIDs[0:namelen]

	return
}

func GetTypeInfoCount(obj interface{}, method uintptr) (c uint32, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		uintptr(2),
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&c)),
		uintptr(0)))
	return
}

func GetTypeInfo(obj interface{}, method uintptr, num uint32) (tinfo *itypeinfo.TypeInfo, err error) {
	err = com.HResultToError(syscall.Syscall6(
		method,
		uintptr(4),
		uintptr(unsafe.Pointer(obj)),
		uintptr(num),
		uintptr(com.GetDefaultUserLocaleID()),
		uintptr(unsafe.Pointer(&tinfo)),
		uintptr(0),
		uintptr(0)))
	return
}

func Invoke(obj interface{}, method uintptr, displayID DisplayID, dispatchContext DispatchContext, params ...interface{}) (result *com.Variant, err error) {
	var displayParams com.DisplayParameter
	var vargs []com.Variant

	if dispatchContext&com.PropertySetDispatchContext != 0 {
		displayNames := [1]int32{com.PropertySetDisplayID}
		displayParams.NamedArgs = uintptr(unsafe.Pointer(&displayNames[0]))
		displayParams.NamedArgsLength = 1
	}

	if len(params) > 0 {
		vargs = make([]com.Variant, len(params))
		for i, v := range params {
			//n := len(params)-i-1
			n := len(params) - i - 1
			vargs[n] = com.VariantByValueType(v)
			// This should clear all variant types.
			defer com.VariantClear(vargs[n])
		}
		displayParams.Args = uintptr(unsafe.Pointer(&vargs[0]))
		displayParams.ArgsLength = uint32(len(params))
	}

	result = new(com.Variant)
	var excepInfo EXCEPINFO
	com.VariantInit(result)
	hr, _, _ := syscall.Syscall9(
		method,
		uintptr(9),
		uintptr(unsafe.Pointer(obj)),
		uintptr(displayID),
		uintptr(unsafe.Pointer(com.NullInterfaceID)),
		uintptr(com.GetDefaultUserLocaleID()),
		uintptr(dispatchContext),
		uintptr(unsafe.Pointer(&displayParams)),
		uintptr(unsafe.Pointer(result)),
		uintptr(unsafe.Pointer(&excepInfo)),
		uintptr(0))
	if hr != 0 {
		err = NewErrorWithSubError(hr, BstrToString(excepInfo.bstrDescription), excepInfo)
	}
	/*
		// This is now deferred and should clear all variants.
		for _, varg := range vargs {
			// This should clear all variant types.
			com.VariantClear(varg)
			if varg.VariantType == BinaryStringVariantType && varg.Val != 0 {
				SysFreeString(((*int16)(unsafe.Pointer(uintptr(varg.Val)))))
			}
			/*
			if varg.VariantType == (BinaryStringVariantType|ByReferenceVariantType) && varg.Val != 0 {
				*(params[n].(*string)) = LpOleStrToString((*uint16)(unsafe.Pointer(uintptr(varg.Val))))
				println(*(params[n].(*string)))
				fmt.Fprintln(os.Stderr, *(params[n].(*string)))
			}
			//* /
		}
	*/
	return
}
