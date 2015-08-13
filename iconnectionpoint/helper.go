package iconnectionpoint

import (
	"unsafe"

	"github.com/go-ole/com"
	"github.com/go-ole/idispatch"
)

// ConnectObject creates a connection point between two services for communication.
func ConnectObject(obj *idispatch.IDispatch, interfaceID *com.GUID, client interface{}) (cookie uint32, err error) {
	var container *ConnectionPointContainer
	err := obj.QueryInterface(InterfaceID, &container)
	if err != nil {
		return
	}
	defer container.Release()

	point, err := container.FindConnectionPoint(interfaceID)
	if err != nil {
		return
	}
	defer point.Release()

	if edisp, ok := client.(*iunknown.Unknown); ok {
		cookie, err = point.Advise(edisp)
		return
	}

	destination, err := com.Service(obj, idispatch.VirtualTableService(), interfaceID)
	if err != nil {
		return
	}

	cookie, err = point.Advise((*iunknown.Unknown)(unsafe.Pointer(destination)))
	if err != nil {
		return
	}

	return
}
