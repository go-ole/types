// +build windows
package iprovideclassinfo

import (
	"unsafe"

	syscall "golang.org/x/sys/windows"

	"github.com/jacobsantos/go-com/pkg/com"
	"github.com/jacobsantos/go-com/pkg/itypeinfo"
)

func GetClassInfo(obj interface{}, method uintptr) (element *itypeinfo.TypeInfo, err error) {
	err = com.HResultToError(syscall.Syscall(
		method,
		2,
		uintptr(unsafe.Pointer(obj)),
		uintptr(unsafe.Pointer(&element)),
		0))
	return
}
